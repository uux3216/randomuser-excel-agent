# Update existing B7344.xlsx — only Work 1 columns B:F

The script only targets `Work 1!B3:F400`.

- B: NAME
- C: BIRTH DAY
- D: ADRESS
- E: PHONE NUMBER
- F: PASSWORD

It does not read, clear, write, add comments to, or format column A or any column from G onward. It also does not edit other worksheets. Data values are stored in Excel comments; cell text is blanked only inside B3:F400.

## GitHub Actions setup

1. Keep `B7344.xlsx` at the repository root.
2. Add/replace `main.py`, `requirements.txt`, and `.github/workflows/update-existing-workbook.yml`.
3. Enable read/write workflow permissions if the workflow needs to push the updated workbook.
4. Open Actions → Update existing Excel comments → Run workflow.
5. The workflow commits the modified `B7344.xlsx` to the same repository.

The API data is fictional test data from RandomUser.me. Make a backup before the first run.
