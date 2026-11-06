JOB ENTRY V2 — INSTALLATION

1. In Apps Script, replace Code.gs with the supplied Code.gs.
2. Save, then Deploy > Manage deployments > Edit (pencil) > select New version > Deploy.
3. Keep the same Web App URL if editing the existing deployment.
4. Upload job-entry.html to the root of your GitHub repository, replacing the previous file.
5. Submit a non-sensitive test record and confirm it appears in columns A:Q.
6. Confirm columns R:S calculate as expected.

Column mapping:
A Designer; B Work Date; C Start; D End; E Job ID; F Customer Name;
G Customer Address; H State; I Mounting Type; J Tool; K Rev #;
L Categories; M Sales Rep; N Remarks; O blank; P Work Type; Q Job Type;
R Duration and S In Hours are preserved/formula columns.

Note: Browser cross-origin submission uses no-cors, so the page cannot read the server response. Verify submissions in the spreadsheet. State detection is based on a US state abbreviation or full state name typed in the address; it is not geocoding.

Updated Web App endpoint to the latest URL supplied by the user.

Confirmed target tab: JOB (gid=885676598).

Final review fix: Start and End Time are converted to numeric Google Sheets time values and formatted as h:mm:ss AM/PM so Duration calculations can use them.
