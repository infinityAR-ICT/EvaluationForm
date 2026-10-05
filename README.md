# EvaluationForm

Evaluation form for teachers. The **Save Responses** button sends all 160 question responses directly to the configured Google Sheet through a Google Apps Script web app. It does not create an XLSX file or use Google Drive OAuth.

## Set up Google Sheets saving

1. Open the destination spreadsheet: [Evaluation responses](https://docs.google.com/spreadsheets/d/12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w/edit).
2. Open **Extensions → Apps Script**. Replace the starter code in `Code.gs` with the contents of this repository's `Code.gs`, then save.
3. In Apps Script, choose **Deploy → New deployment → Web app**.
4. Set **Execute as** to your account and **Who has access** to **Anyone**, then deploy and complete Google's authorization prompts. This lets respondents submit without signing in; the script writes using the deploying account's spreadsheet access.
5. Copy the deployed web app URL ending in `/exec`. In `index.html`, replace `YOUR_APPS_SCRIPT_WEB_APP_URL` in `GOOGLE_SHEETS_WEB_APP_URL` with that URL. Publish the updated form.
6. Test with a submission, then check the spreadsheet's `Responses` tab. The script creates the tab and header row if needed, then appends one row per question, grouped by submission ID and timestamp.

The public web app URL is not authentication. Anyone who obtains it can submit rows to the spreadsheet as the deploying account, potentially causing unwanted submissions or using your Apps Script quota. Share the form and endpoint carefully, monitor the sheet, and disable or redeploy the web app if it is abused. An anonymous endpoint cannot provide the protections of user sign-in.

The browser uses a cross-origin `no-cors` request because Apps Script web apps do not expose their response to this page. The form reports that the request was sent, but cannot confirm the server-side result; verify the `Responses` tab if a submission does not appear.
