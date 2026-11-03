# RandomUser Excel Agent — values stored in Comments

This GitHub Actions project fetches fictional test profiles from the RandomUser.me API and creates `users.xlsx`.

## Important: data is in Excel comments, not cell text

- Header cells contain `Name`, `Address`, `Phone`, `Email`, `ID`.
- Each person occupies one row.
- The data cells are blank.
- Each data value is stored in that cell's Excel comment (legacy comment/note), authored by `RandomUser Agent`.
- Select or tap a data cell to view its comment.

## Run with GitHub Actions

1. Create a GitHub repository and upload all project files, preserving `.github/workflows/generate-excel.yml`.
2. Open **Actions** → **Generate Random User Excel** → **Run workflow**.
3. Set record count to `500` and seed to `excel-agent-demo`.
4. After completion, download the `randomuser-excel` artifact from the workflow run. It contains `users.xlsx`.

## Columns

Name, Address, Phone, Email, ID. ID uses `login.uuid` with the API `id.value` as a fallback.

## Local run (optional)

```bash
python -m pip install -r requirements.txt
python main.py
```

The default is 500 records. Set `RECORD_COUNT` to a value between 1 and 5000 to change it.

## Notes

- Uses the official RandomUser.me API, not browser clicking.
- These are fictional test profiles; do not use their contact fields to contact or identify real people.
- Excel comments are supported by `openpyxl`. They are conventional Excel comments/notes, not Microsoft 365 threaded comments.
- The workflow runs manually and stores the artifact for 14 days.
