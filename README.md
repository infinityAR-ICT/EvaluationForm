# EvaluationForm

Evaluation form for teachers. The **Save Responses** button sends all 120 question responses directly to the configured Google Sheet through a Google Apps Script web app. It does not create an XLSX file or use Google Drive OAuth.

## Set up Google Sheets saving

1. Open the destination spreadsheet: [Evaluation responses](https://docs.google.com/spreadsheets/d/12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w/edit).
2. Open **Extensions → Apps Script**. Replace the script code with the contents of this repository's `Code.gs`, then save.
3. In the Apps Script editor, select and run `renameSectionSheets` once to rename existing `Section 1` through `Section 30` tabs to the corresponding photo filenames (without the file extension or leading number). Approve the authorization prompt. The script also performs these renames automatically when it receives a submission.
4. In Apps Script, choose **Deploy → Manage deployments**. Edit the web app deployment, choose **New version**, and deploy it. If you have not deployed it before, choose **Deploy → New deployment → Web app** instead.
5. Set **Execute as** to your account and **Who has access** to **Anyone**, then deploy and complete Google's authorization prompts. This lets respondents submit without signing in; the script writes using the deploying account's spreadsheet access.
6. Keep the existing `/exec` web app URL if updating a deployment. If creating a new deployment, copy its `/exec` URL and update `GOOGLE_SHEETS_WEB_APP_URL` in `index.html`. Publish the updated form if you changed the URL.
7. Submit a response and check the spreadsheet. The selected class, house, and form, along with the optional cadet name and number, are recorded in the final `Class`, `House`, `Form`, `Cadet Name`, and `Cadet Number` columns on each person tab. The section title is merged across D2:I2, headers are on row 4 (`No`, `Submit on`, `Id`, then each question's `Response` and `Comment` columns, followed by those identity columns), and each submission adds exactly one row per person tab starting at row 5.

Each row contains a per-section response number, timestamp, submission ID, and the four selected options and comments for that section. The current form has 30 photo sections and 120 questions. Deploy the updated Apps Script code before testing; an old deployment will keep writing in its old format.

Older rows in the previous `Responses` tab are not moved or deleted. They can be removed manually after confirming you no longer need them.

The public web app URL is not authentication. Anyone who obtains it can submit rows to the spreadsheet as the deploying account, potentially causing unwanted submissions or using your Apps Script quota. Share the form and endpoint carefully, monitor the sheet, and disable or redeploy the web app if it is abused. An anonymous endpoint cannot provide the protections of user sign-in.

The browser uses a cross-origin `no-cors` request because Apps Script web apps do not expose their response to this page. The form reports that the request was sent, but cannot confirm the server-side result; verify the `Responses` tab if a submission does not appear.
