const SPREADSHEET_ID = "12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w";
const RESPONSE_SHEET_NAME = "Responses";
const OPTIONS = ["Option 1", "Option 2", "Option 3", "Option 4"];
const SECTION_COUNT = 40;
const QUESTIONS_PER_SECTION = 4;
const MAX_COMMENT_LENGTH = 45000;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const answers = validateAnswers(payload.answers);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(RESPONSE_SHEET_NAME)
      || spreadsheet.insertSheet(RESPONSE_SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Submitted At",
        "Submission ID",
        "Section",
        "Question Number",
        "Question",
        "Selected Option",
        "Comment"
      ]);
    }

    const submittedAt = new Date();
    const submissionId = Utilities.getUuid();
    const rows = answers.map(answer => [
      submittedAt,
      submissionId,
      answer.section,
      answer.question,
      safeCellText(answer.questionText),
      answer.selectedOption,
      safeCellText(answer.comment)
    ]);

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
    return jsonResponse({ok: true, submissionId: submissionId, rowsAdded: rows.length});
  } catch (error) {
    console.error(error);
    return jsonResponse({ok: false, error: "The response could not be saved."});
  } finally {
    lock.releaseLock();
  }
}

function validateAnswers(answers) {
  const expectedCount = SECTION_COUNT * QUESTIONS_PER_SECTION;
  if (!Array.isArray(answers) || answers.length !== expectedCount) {
    throw new Error("Invalid answer count.");
  }

  return answers.map((answer, index) => {
    const expectedSection = Math.floor(index / QUESTIONS_PER_SECTION) + 1;
    const expectedQuestion = (index % QUESTIONS_PER_SECTION) + 1;
    if (!answer
      || answer.section !== expectedSection
      || answer.question !== expectedQuestion
      || typeof answer.questionText !== "string"
      || typeof answer.comment !== "string"
      || answer.comment.length > MAX_COMMENT_LENGTH
      || (answer.selectedOption !== "" && !OPTIONS.includes(answer.selectedOption))) {
      throw new Error(`Invalid answer at row ${index + 1}.`);
    }
    return answer;
  });
}

function safeCellText(value) {
  return /^[\s]*[=+\-@]/.test(value) ? `'${value}` : value;
}

function jsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
