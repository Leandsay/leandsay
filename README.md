
# Job Entry Form

A responsive standalone form UI backed by Google Apps Script. Submissions append a row to the Google Sheet using the supplied A:P column layout.

## Files
- `Code.gs` — Apps Script server-side functions.
- `Index.html` — Form UI.

## Important: GitHub vs. the Live Form

GitHub is used to store and version the source files. The form must be deployed as a **Google Apps Script Web App** because it uses `google.script.run` to securely call the Apps Script functions that write to Google Sheets.

Do not deploy `Index.html` as a GitHub Pages site and expect it to write to Sheets. `google.script.run` is only available inside Apps Script HTML Service.

## Setup

1. Open the Google spreadsheet containing the existing Aurora log and dashboard.
2. Go to **Extensions → Apps Script**.
3. Replace the default script with the contents of `Code.gs`.
4. Click **+ → HTML**, name the file exactly `Index`, and paste the contents of `Index.html`.
5. Save the project.
6. In Apps Script, select **Deploy → New deployment**.
7. Select the deployment type **Web app**.
8. Set **Execute as** to **Me**. Choose who can access it according to your company policy.
9. Click **Deploy**, authorize the requested permissions, and copy the Web app URL. This is the shareable form link.

## Sheet Layout Expected

The script writes columns A:P:

| Column | Field |
|---|---|
| A | Designers |
| B | Work Date |
| C | Start |
| D | End |
| E | Job ID |
| F | Customer Name |
| G | Customer Address |
| H | State |
| I | M-Type |
| J | Tool |
| K | Rev # |
| L | Categories |
| M | Sales Rep |
| N | Remarks |
| O | Reserved blank column |
| P | Work Type |

The script searches the first 10 rows for the expected headers. To target one tab explicitly, set `SHEET_NAME` in `CONFIG` inside `Code.gs` to the exact tab name.

## Dropdowns

The form tries to reuse data-validation options from the first data row of the corresponding columns. If no validation rule is present, it builds dropdown choices from existing non-empty values in that column.

Keep the dropdown validations or values in the destination sheet maintained as your team changes.

## Notes

- `JOB` entries require Job ID and Customer Name.
- `BREAK` and `DOWNTIME` entries do not require Job ID or Customer Name.
- Column O is intentionally left blank to preserve the supplied sheet layout.
- The form appends rows; it does not overwrite existing entries.
- The Apps Script project must be bound to the destination spreadsheet.
