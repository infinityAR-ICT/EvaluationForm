const COLLEGE_SPREADSHEET_ID = "12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w";
const COLLEGE_SHEET_NAME = "College Evaluation";
const COLLEGE_OPTIONS = ["Excellent", "Standard", "Good", "Satisfactory", "Poor"];
const COLLEGE_MAX_COMMENT_LENGTH = 45000;
const COLLEGE_VALID_CLASSES = ["VII", "VIII", "IX", "X", "XI", "XII"];
const COLLEGE_VALID_HOUSES = [
  "Shaheed Suhrawardy House",
  "Sher-e-Bangla House",
  "Shariatullah House"
];
const COLLEGE_VALID_FORMS = ["A", "B"];
const COLLEGE_RATING_FIELDS = [
  {key: "environmentRating", header: "College Environment Rating"},
  {key: "educationDevelopmentRating", header: "Education Development Measures Rating"},
  {key: "diningQualityRating", header: "Cadet Dining and Food Quality Rating"},
  {key: "houseEnvironmentRating", header: "House Environment Rating"},
  {key: "formMasterSincerityRating", header: "Form Master Sincerity Rating"},
  {key: "houseMasterDedicationRating", header: "House Master Sincerity and Dedication Rating"},
  {key: "hospitalFacilitiesRating", header: "Hospital Facilities and Environment Rating"},
  {key: "collegeAdjutantSincerityRating", header: "College Adjutant Sincerity Rating"}
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!e || !e.postData || typeof e.postData.contents !== "string") {
      throw new Error("Missing request body.");
    }
    const payload = JSON.parse(e.postData.contents);
    if (payload.evaluationType && payload.evaluationType !== "college") {
      throw new Error("Invalid evaluation type.");
    }
    const submission = {
      className: collegeValidateChoice(payload.className, COLLEGE_VALID_CLASSES, "class"),
      houseName: collegeValidateChoice(payload.houseName, COLLEGE_VALID_HOUSES, "house"),
      formName: collegeValidateChoice(payload.formName, COLLEGE_VALID_FORMS, "form"),
      cadetName: collegeValidateCadetName(payload.cadetName),
      cadetNumber: collegeValidateCadetNumber(payload.cadetNumber),
      answers: collegeValidateAnswers(payload.collegeAnswers),
      submittedAt: new Date()
    };

    const spreadsheet = SpreadsheetApp.openById(COLLEGE_SPREADSHEET_ID);
    collegeSaveSubmission(spreadsheet, submission);
    SpreadsheetApp.flush();
    return collegeJsonResponse({ok: true, rowsAdded: 1});
  } catch (error) {
    console.error(error);
    return collegeJsonResponse({ok: false, error: "The response could not be saved."});
  } finally {
    lock.releaseLock();
  }
}

function collegeSaveSubmission(spreadsheet, submission) {
  let sheet = spreadsheet.getSheetByName(COLLEGE_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(COLLEGE_SHEET_NAME);
  }

  const headers = collegeBuildHeaders();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(2, 9, 1, 6).merge();
    sheet.getRange(2, 9).setValue(COLLEGE_SHEET_NAME);
    sheet.getRange(4, 2, 1, headers.length).setValues([headers]);
  } else {
    collegeMigrateOrValidateHeaders(sheet, headers);
    sheet.getRange(2, 9).setValue(COLLEGE_SHEET_NAME);
  }

  sheet.getRange(2, 9, 1, 6)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.getRange(4, 2, 1, headers.length)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setFrozenRows(4);
  sheet.setIndex(spreadsheet.getSheets().length);

  const rowNumber = collegeGetNextResponseRow(sheet);
  const row = [
    rowNumber - 4,
    submission.submittedAt,
    submission.className,
    submission.houseName,
    submission.formName,
    collegeSafeCellText(submission.cadetName),
    collegeSafeCellText(submission.cadetNumber),
    ...COLLEGE_RATING_FIELDS.flatMap(field => [
      submission.answers.ratings[field.key],
      collegeSafeCellText(submission.answers.comments[field.key])
    ]),
    collegeSafeCellText(submission.answers.missionDescription)
  ];
  sheet.getRange(rowNumber, 2, 1, row.length).setValues([row]);
}

function collegeBuildHeaders() {
  const headers = [
    "No",
    "Submit on",
    "Class",
    "House",
    "Form",
    "Cadet Name",
    "Cadet Number"
  ];
  COLLEGE_RATING_FIELDS.forEach(field => headers.push(field.header, `${field.header} Comment`));
  headers.push("College Mission or Purpose");
  return headers;
}

function collegeMigrateOrValidateHeaders(sheet, headers) {
  if (sheet.getLastRow() < 4) {
    throw new Error("The existing College Evaluation sheet has an incomplete layout.");
  }

  const oldHeaders = sheet.getRange(4, 2, 1, headers.length).getValues()[0];
  if (headers.every((header, index) => oldHeaders[index] === header)) {
    return;
  }

  const basicHeaders = ["No", "Submit on", "Class", "House", "Form", "Cadet Name", "Cadet Number"];
  const hasBasicHeaders = basicHeaders.every((header, index) => oldHeaders[index] === header);
  const ratingOnlyLayout = hasBasicHeaders
    && oldHeaders[7] === COLLEGE_RATING_FIELDS[0].header
    && oldHeaders[8] === "College Mission or Purpose";
  const fullLayout = hasBasicHeaders
    && COLLEGE_RATING_FIELDS.every((field, index) => oldHeaders[index + 7] === field.header)
    && oldHeaders[15] === "College Mission or Purpose";

  if (!ratingOnlyLayout && !fullLayout) {
    throw new Error("The existing College Evaluation sheet has incompatible headers.");
  }

  const lastRow = sheet.getLastRow();
  const oldColumnCount = fullLayout ? 16 : 9;
  const oldRows = sheet.getRange(4, 2, lastRow - 3, oldColumnCount).getValues();
  const migratedRows = oldRows.map((oldRow, rowIndex) => {
    if (rowIndex === 0) {
      return headers;
    }
    const ratings = fullLayout
      ? oldRow.slice(7, 15)
      : [oldRow[7], ...Array(COLLEGE_RATING_FIELDS.length - 1).fill("")];
    return [
      ...oldRow.slice(0, 7),
      ...ratings.flatMap(rating => [rating, ""]),
      fullLayout ? oldRow[15] : oldRow[8]
    ];
  });
  sheet.getRange(4, 2, migratedRows.length, headers.length).setValues(migratedRows);
}

function collegeValidateAnswers(answers) {
  if (!answers
    || !answers.ratings
    || !answers.comments
    || typeof answers.missionDescription !== "string"
    || !answers.missionDescription.trim()
    || answers.missionDescription.length > 5000) {
    throw new Error("Invalid or incomplete College Evaluation answers.");
  }

  const ratings = {};
  const comments = {};
  COLLEGE_RATING_FIELDS.forEach(field => {
    const rating = answers.ratings[field.key];
    const comment = answers.comments[field.key];
    if (!COLLEGE_OPTIONS.includes(rating)) {
      throw new Error(`Invalid or missing ${field.header}.`);
    }
    if (typeof comment !== "string"
      || comment.length > COLLEGE_MAX_COMMENT_LENGTH
      || ((rating === "Satisfactory" || rating === "Poor") && !comment.trim())) {
      throw new Error(`Invalid or missing comment for ${field.header}.`);
    }
    ratings[field.key] = rating;
    comments[field.key] = comment.trim();
  });

  return {
    ratings: ratings,
    comments: comments,
    missionDescription: answers.missionDescription.trim()
  };
}

function collegeValidateChoice(value, allowedValues, fieldName) {
  if (!allowedValues.includes(value)) {
    throw new Error(`Invalid or missing ${fieldName}.`);
  }
  return value;
}

function collegeValidateOptionalText(value, fieldName, maxLength) {
  if (value === undefined || value === null || value === "") {
    return "";
  }
  if (typeof value !== "string" || value.length > maxLength) {
    throw new Error(`Invalid ${fieldName}.`);
  }
  return value.trim();
}

function collegeValidateCadetName(value) {
  const name = collegeValidateOptionalText(value, "Cadet Name", 200);
  if (name && !/^[\p{L}\p{M}]+(?:\s+[\p{L}\p{M}]+)*$/u.test(name)) {
    throw new Error("Cadet Name may contain letters and spaces only.");
  }
  return name;
}

function collegeValidateCadetNumber(value) {
  const number = collegeValidateOptionalText(value, "Cadet Number", 100);
  if (number && !/^\d+$/.test(number)) {
    throw new Error("Cadet Number may contain digits only.");
  }
  return number;
}

function collegeGetNextResponseRow(sheet) {
  const firstDataRow = 5;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) {
    return firstDataRow;
  }

  const numbers = sheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, 1).getValues();
  for (let index = numbers.length - 1; index >= 0; index--) {
    if (numbers[index][0] !== "" && numbers[index][0] !== null) {
      return firstDataRow + index + 1;
    }
  }
  return firstDataRow;
}

function collegeSafeCellText(value) {
  return /^[\s]*[=+\-@]/.test(value) ? `'${value}` : value;
}

function collegeJsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
