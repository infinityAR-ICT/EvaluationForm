const SPREADSHEET_ID = "12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w";
const OPTIONS = ["Option 1", "Option 2", "Option 3", "Option 4"];
const SECTION_COUNT = 30;
const QUESTIONS_PER_SECTION = 4;
const MAX_COMMENT_LENGTH = 45000;
const VALID_CLASSES = ["VII", "VIII", "IX", "X", "XI", "XII"];
const VALID_HOUSES = [
  "Shaheed Suhrawardy House",
  "Sher-e-Bangla House",
  "Shariatullah House"
];
const VALID_FORMS = ["A", "B"];
const SECTION_PHOTO_FILES = [
  "1. MRS. JAMUNA RANI BISWAS.jpg",
  "2. MD. TAUFIQUL ALAM.jpg",
  "3. MD. TAREEKUL HAQ.jpg",
  "4. MR. PRODIPTA KUMAR.jpg",
  "5. MD. MAHBUBUL ALAM.jpg",
  "6. MRS. ASMA-UL- MAHMUDA TASLIMA.jpg",
  "7. MD. MAIN UDDIN KHAN.JPG",
  "8. MR. MUHAMMAD SHAHAB UDDIN.jpg",
  "9. MRS. MUKTI RANI MODAK.JPG",
  "10. MR. SYED SELIMUZZAMAN.JPG",
  "11. MRS. SHAFINA RAHAT .jpg",
  "12. MUHAMMD ABUL KALAM AZAD.jpg",
  "13. MD. MOHIUDDIN KHAN.jpg",
  "14. MD. ARIF KHAN .jpg",
  "15.  MD. NAZMUS SHAHID.jpg",
  "16. MISS MUNNI.jpg",
  "17. MD. NAZRUL ISLAM.jpg",
  "18. MD. MATIUR RAHMAN.jpg",
  "19. MR. INDRAJIT KUNDU.jpg",
  "20. MR. ABU NAYEEM MOHAMMAD EKRAM.JPG",
  "21. MR. JOYDEV MONDAL.JPG",
  "22. MR. METUN MONDOL.jpg",
  "23. MD. ABDUL MOMIN.jpg",
  "24. MR. MOHSIN EMRAN.jpg",
  "25. MR.NAZMUL HASAN.jpg",
  "26. MR. SAZZADUR RAHMAN .jpg",
  "27. MR.  NAZIM AL HASAN.jpg",
  "28. MRS. TASKIA.jpg",
  "29. MST. JANNATUN  FERDOUS.jpg",
  "30. MD. IBRAHIM MOLLA.jpg"
];

function getSectionSheetNames() {
  const usedNames = {};
  return SECTION_PHOTO_FILES.map(fileName => {
    const baseName = fileName.replace(/^\d+\.\s*/, "").replace(/\.[^.]+$/, "").trim();
    let sheetName = baseName;
    let suffix = 2;
    while (usedNames[sheetName.toLowerCase()]) {
      sheetName = `${baseName} (${suffix})`;
      suffix++;
    }
    usedNames[sheetName.toLowerCase()] = true;
    return sheetName;
  });
}

function renameSectionSheets() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetNames = getSectionSheetNames();
  const renames = [];

  for (let section = 1; section <= SECTION_COUNT; section++) {
    const oldSheet = spreadsheet.getSheetByName(`Section ${section}`);
    const targetSheet = spreadsheet.getSheetByName(sheetNames[section - 1]);
    if (oldSheet && targetSheet && oldSheet.getSheetId() !== targetSheet.getSheetId()) {
      throw new Error(`Both "Section ${section}" and "${sheetNames[section - 1]}" exist. Resolve the duplicate tabs before renaming.`);
    }
    if (oldSheet && !targetSheet) {
      renames.push({sheet: oldSheet, name: sheetNames[section - 1]});
    }
  }

  renames.forEach(rename => {
    rename.sheet.setName(rename.name);
    rename.sheet.getRange(2, 4).setValue(rename.name);
  });
  SpreadsheetApp.flush();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const className = validateClassName(payload.className);
    const houseName = validateHouseName(payload.houseName);
    const formName = validateFormName(payload.formName);
    const cadetName = validateOptionalText(payload.cadetName, "Cadet Name", 200);
    const cadetNumber = validateOptionalText(payload.cadetNumber, "Cadet Number", 100);
    const answers = validateAnswers(payload.answers);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const submittedAt = new Date();
    const submissionId = Utilities.getUuid();
    const answersBySection = Array.from({length: SECTION_COUNT}, () => []);
    answers.forEach(answer => answersBySection[answer.section - 1].push(answer));
    const sections = [];
    const sheetNames = getSectionSheetNames();

    for (let section = 1; section <= SECTION_COUNT; section++) {
      const sheetName = sheetNames[section - 1];
      let sheet = spreadsheet.getSheetByName(sheetName);
      const oldSheet = spreadsheet.getSheetByName(`Section ${section}`);
      if (sheet && oldSheet && sheet.getSheetId() !== oldSheet.getSheetId()) {
        throw new Error(`Both "Section ${section}" and "${sheetName}" exist. Resolve the duplicate tabs before submitting.`);
      }
      if (!sheet && oldSheet) {
        oldSheet.setName(sheetName);
        sheet = oldSheet;
      }
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
      }
      const headers = buildHeaders();
      const isEmpty = sheet.getLastRow() === 0;

      if (!isEmpty) {
        const responseHeaders = headers.slice(0, -5);
        const existingHeaders = sheet.getRange(4, 1, 1, responseHeaders.length).getValues()[0];
        if (responseHeaders.some((header, index) => existingHeaders[index] !== header)) {
          throw new Error(`Unexpected headers in ${sheetName}.`);
        }
        for (let index = responseHeaders.length; index < headers.length; index++) {
          const existingHeader = sheet.getRange(4, index + 1).getValue();
          if (existingHeader && existingHeader !== headers[index]) {
            throw new Error(`Unexpected ${headers[index]} column header in ${sheetName}.`);
          }
        }
      }

      sections.push({
        sheet: sheet,
        sectionNumber: section,
        sheetName: sheetName,
        headers: headers,
        isEmpty: isEmpty
      });
    }

    sections.forEach(section => {
      const sheet = section.sheet;
      if (section.isEmpty) {
        sheet.getRange(2, 4, 1, 6).merge();
        sheet.getRange(4, 1, 1, section.headers.length).setValues([section.headers]);
      } else {
        const responseHeaderCount = section.headers.length - 5;
        for (let index = responseHeaderCount; index < section.headers.length; index++) {
          if (!sheet.getRange(4, index + 1).getValue()) {
            sheet.getRange(4, index + 1).setValue(section.headers[index]);
          }
        }
      }
      sheet.getRange(2, 4).setValue(section.sheetName);
      sheet.getRange(2, 4, 1, 6)
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      sheet.getRange(4, 1, 1, section.headers.length)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
      sheet.setFrozenRows(4);

      const rowNumber = sheet.getLastRow() + 1;
      const row = [rowNumber - 4, submittedAt, submissionId];
      answersBySection[section.sectionNumber - 1].forEach(answer => {
        row.push(answer.selectedOption, safeCellText(answer.comment));
      });
      row.push(className);
      row.push(houseName);
      row.push(formName);
      row.push(safeCellText(cadetName));
      row.push(safeCellText(cadetNumber));
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    });

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
  const headers = ["No", "Submit on", "Id"];
  for (let question = 1; question <= QUESTIONS_PER_SECTION; question++) {
    headers.push(`Question ${question}: Response`, "Comment");
  }
  headers.push("Class");
  headers.push("House");
  headers.push("Form");
  headers.push("Cadet Name");
  headers.push("Cadet Number");
  return headers;
}

function validateClassName(className) {
  if (!VALID_CLASSES.includes(className)) {
    throw new Error("Invalid or missing class.");
  }
  return className;
}

function validateHouseName(houseName) {
  if (!VALID_HOUSES.includes(houseName)) {
    throw new Error("Invalid or missing house.");
  }
  return houseName;
}

function validateFormName(formName) {
  if (!VALID_FORMS.includes(formName)) {
    throw new Error("Invalid or missing form.");
  }
  return formName;
}

function validateOptionalText(value, fieldName, maxLength) {
  if (value === undefined || value === null || value === "") {
    return "";
  }
  if (typeof value !== "string" || value.length > maxLength) {
    throw new Error(`Invalid ${fieldName}.`);
  }
  return value.trim();
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
