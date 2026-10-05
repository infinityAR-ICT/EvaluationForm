# EvaluationForm

Evaluation form for teachers. The form can export the answers to an Excel workbook or save the workbook to a configured Google Drive folder.

## Configure Google Drive saving

Google Drive saving uses Google sign-in, Google Picker, and the Drive API's limited `drive.file` permission. It does not use a backend or grant the page general access to a user's Drive.

1. In Google Cloud Console, create or select a project and enable **Google Drive API** and **Google Picker API**.
2. Configure the OAuth consent screen. Add test users while testing; Google may require additional consent-screen setup or verification before public use.
3. Create an OAuth client ID of type **Web application**. Add the exact site origin (scheme + host, without a path) under **Authorized JavaScript origins**.
4. Create an API key for Google Picker. Restrict it to the deployed site's HTTP referrers and the Google Picker API.
5. In `index.html`, replace `YOUR_GOOGLE_OAUTH_CLIENT_ID` and `YOUR_GOOGLE_API_KEY` with the values from Google Cloud Console. The API key is visible in the page, so keep its referrer/API restrictions enabled.
6. The destination folder ID is already set to `15Y5Dmle5QvCul7KGRbE3RbXMSEcRCDQZ`. To change the destination, update `GOOGLE_DRIVE_FOLDER_ID` to the folder ID from its Drive URL.
7. Give each user Google account **Editor** access to the destination folder. Users sign in, select that exact folder in Google Picker the first time, and then the form saves subsequent workbooks there. The selected folder ID is stored in that browser's local storage.

The **Save / Export to Excel** button saves the workbook to Google Drive, and the final-section **Finish** button does the same. If Drive is not configured or upload fails, the form displays an error.
