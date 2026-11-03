# RandomUser → Google Sheets native comment threads

This workflow creates a **new native Google Spreadsheet** (not an `.xlsx` file) with 398 fictional RandomUser profiles.

## What it creates

- Sheet: `Generated Data`
- Exact cell range used: `A1:E398`
- Every cell value is `Entry`
- Every cell is bold, italic, font size 11
- One native Google Sheets comment thread per cell (1,990 total)
- Comment content by column: A name, B birthday, C address, D phone, E password
- No replies are created
- The script makes no write, formatting, or comment requests to column F or later.

This uses the Google Sheets API's `insertComment` request. It does not build comments by writing raw OOXML, and it does not create Excel Notes/legacy comments.

## One-time setup

This setup is required because GitHub Actions must be authorized to create a Google Sheet in Google Drive.

1. Open Google Cloud Console and create/select a project.
2. Enable **Google Sheets API** and **Google Drive API**.
3. Create a service account and download its JSON key.
4. In GitHub, open the repository → **Settings → Secrets and variables → Actions**.
5. Add these repository secrets:
   - `GOOGLE_SERVICE_ACCOUNT_JSON`: paste the *entire contents* of the service-account JSON key file.
   - `SHARE_WITH_EMAIL`: your Google account email address. The created sheet will be shared with this account as Editor.
6. Make sure the service account is allowed to create files in Drive under the project's configuration. Some Google Workspace accounts restrict service accounts; if so, use an OAuth user-token workflow instead.
7. Go to **Actions → Create Google Sheets with native comment threads → Run workflow**.

## Where the result appears

The workflow log prints `SPREADSHEET_URL`. Open the run log and tap/copy that URL. The sheet is also shared with the email in `SHARE_WITH_EMAIL`.

## Important notes

- The profiles come from RandomUser and are fictional test data. Do not use generated passwords for real accounts.
- The generated file is a Google Spreadsheet. This is the right format to verify native Google Sheets comment threads in the Google Sheets app.
- Exporting to `.xlsx` may not preserve all Google Sheets-native comment behavior exactly; verify in Google Sheets first.
- The script verifies the comment count using the Sheets API and fails if it cannot read back all 1,990 comment threads.
- If the API returns an error, read the GitHub Actions log. Google APIs/permissions can vary by account and project.
