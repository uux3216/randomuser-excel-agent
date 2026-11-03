import os
import sys
import requests
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

API_URL = "https://randomuser.me/api/1.4/"
COUNT = int(os.getenv("RECORD_COUNT", "500"))
SEED = os.getenv("RANDOMUSER_SEED", "excel-agent-demo")
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "users.xlsx")
COMMENT_AUTHOR = "RandomUser Agent"

def full_name(user):
    name = user.get("name") or {}
    return " ".join(part for part in [name.get("title"), name.get("first"), name.get("last")] if part).strip()

def full_address(user):
    loc = user.get("location") or {}
    street = loc.get("street") or {}
    street_text = " ".join(str(street.get(k, "")).strip() for k in ("number", "name") if street.get(k) not in (None, ""))
    parts = [street_text, loc.get("city"), loc.get("state"), loc.get("postcode"), loc.get("country")]
    return ", ".join(str(part).strip() for part in parts if part not in (None, ""))

def fetch_users(count, seed):
    if not 1 <= count <= 5000:
        raise ValueError("RECORD_COUNT must be between 1 and 5000.")
    params = {
        "results": count,
        "seed": seed,
        "inc": "name,location,email,phone,login,id",
        "noinfo": "1",
    }
    response = requests.get(API_URL, params=params, timeout=60)
    response.raise_for_status()
    payload = response.json()
    if isinstance(payload, dict) and payload.get("error"):
        raise RuntimeError(f"RandomUser API error: {payload['error']}")
    users = payload.get("results")
    if not isinstance(users, list):
        raise RuntimeError("API response did not contain a results list.")
    if len(users) != count:
        raise RuntimeError(f"Expected {count} records, received {len(users)}.")
    return users

def add_comment(cell, value):
    # Leave the cell's displayed value blank; store the data only in its Excel comment.
    cell.value = None
    cell.comment = Comment(str(value) if value not in (None, "") else "", COMMENT_AUTHOR)

def write_excel(users, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Users"
    headers = ["Name", "Address", "Phone", "Email", "ID"]
    ws.append(headers)

    for user in users:
        login = user.get("login") or {}
        generated_id = login.get("uuid") or ((user.get("id") or {}).get("value")) or ""
        values = [
            full_name(user),
            full_address(user),
            str(user.get("phone") or ""),
            str(user.get("email") or ""),
            str(generated_id),
        ]
        ws.append([None] * len(headers))
        row = ws.max_row
        for col, value in enumerate(values, start=1):
            add_comment(ws.cell(row=row, column=col), value)

    header_fill = PatternFill("solid", fgColor="145DA0")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{ws.max_row}"

    widths = [28, 52, 22, 34, 38]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(output_file)
    print(f"SUCCESS: saved {len(users)} fictional test records to {output_file}")
    print("Each data value is stored in its cell's Excel comment; data cells are blank.")

def main():
    try:
        print(f"Fetching {COUNT} fictional test records from RandomUser.me...")
        users = fetch_users(COUNT, SEED)
        write_excel(users, OUTPUT_FILE)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise

if __name__ == "__main__":
    main()
