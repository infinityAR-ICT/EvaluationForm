const SPREADSHEET_ID = "12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w";
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
    const submittedAt = new Date();
    const submissionId = Utilities.getUuid();
    const answersBySection = Array.from({length: SECTION_COUNT}, () => []);
    answers.forEach(answer => answersBySection[answer.section - 1].push(answer));

    for (let section = 1; section <= SECTION_COUNT; section++) {
      const sheetName = `Section ${section}`;
      const sheet = spreadsheet.getSheetByName(sheetName)
        || spreadsheet.insertSheet(sheetName);
      const headers = buildHeaders();

      if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
      } else {
        const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
        if (headers.some((header, index) => existingHeaders[index] !== header)) {
          throw new Error(`Unexpected headers in ${sheetName}.`);
        }
      }

      const row = [submittedAt, submissionId];
      answersBySection[section - 1].forEach(answer => {
        row.push(answer.selectedOption, safeCellText(answer.comment));
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    }

    SpreadsheetApp.flush();
    return jsonResponse({
      ok: true,
      submissionId: submissionId,
      sectionsUpdated: SECTION_COUNT,
      rowsAdded: SECTION_COUNT
    });
  } catch (error) {
    console.error(error);
    return jsonResponse({ok: false, error: "The response could not be saved."});
  } finally {
    lock.releaseLock();
  }
}

function buildHeaders() {
  const headers = ["Submitted At", "Submission ID"];
  for (let question = 1; question <= QUESTIONS_PER_SECTION; question++) {
    headers.push(`Question ${question} - Selected Option`, `Question ${question} - Comment`);
  }
  return headers;
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
