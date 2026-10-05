const SPREADSHEET_ID = "12UGu3FmqVhGbjwCyJUiR4owIqwU6klFNB0yzU1SRH-w";
const OPTIONS = ["Excellent", "Standard", "Good", "Satisfactory", "Poor"];
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
const QUESTION_HEADERS = [
  "Subject Knowledge",
  "Teaching Skill",
  "Sincerity towards The Cadets",
  "Personality & Behaviour"
];
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

function archiveSheet(spreadsheet, sheet, originalName) {
  const baseName = `${originalName} - Legacy Archive`;
  let archiveName = baseName;
  let suffix = 2;
  while (spreadsheet.getSheetByName(archiveName)) {
    archiveName = `${baseName} (${suffix})`;
    suffix++;
  }
  sheet.setName(archiveName);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const className = validateClassName(payload.className);
    const houseName = validateHouseName(payload.houseName);
    const formName = validateFormName(payload.formName);
    const cadetName = validateCadetName(payload.cadetName);
    const cadetNumber = validateCadetNumber(payload.cadetNumber);
    const evaluationType = payload.evaluationType || "teacher";
    if (evaluationType !== "teacher") {
      throw new Error("Invalid evaluation type.");
    }
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const submittedAt = new Date();
    const submissionId = Utilities.getUuid();

    const answers = validateAnswers(payload.answers);
    const answersBySection = Array.from({length: SECTION_COUNT}, () => []);
    answers.forEach(answer => answersBySection[answer.section - 1].push(answer));
    const sections = [];
    const sheetNames = getSectionSheetNames();

    for (let section = 1; section <= SECTION_COUNT; section++) {
      const sheetName = sheetNames[section - 1];
      let sheet = spreadsheet.getSheetByName(sheetName);
      const oldSheet = spreadsheet.getSheetByName(`Section ${section}`);
      if (sheet && oldSheet && sheet.getSheetId() !== oldSheet.getSheetId()) {
        archiveSheet(spreadsheet, oldSheet, `Section ${section}`);
      }
      if (!sheet && oldSheet) {
        oldSheet.setName(sheetName);
        sheet = oldSheet;
      }
      const headers = buildHeaders();
      if (sheet && sheet.getLastRow() > 0) {
        const existingHeaders = sheet.getRange(4, 2, 1, headers.length).getValues()[0];
        if (headers.some((header, index) => existingHeaders[index] !== header)) {
          archiveSheet(spreadsheet, sheet, sheetName);
          sheet = null;
        }
      }
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
      }
      const isEmpty = sheet.getLastRow() === 0;

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
        sheet.getRange(2, 9, 1, 6).merge();
        sheet.getRange(4, 2, 1, section.headers.length).setValues([section.headers]);
      }
      sheet.getRange(2, 9).setValue(section.sheetName);
      sheet.getRange(2, 9, 1, 6)
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      sheet.getRange(4, 2, 1, section.headers.length)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
      sheet.setFrozenRows(4);

      const rowNumber = getNextResponseRow(sheet);
      const row = [rowNumber - 4, submittedAt, className, houseName, formName,
        safeCellText(cadetName), safeCellText(cadetNumber)];
      answersBySection[section.sectionNumber - 1].forEach(answer => {
        row.push(answer.selectedOption, safeCellText(answer.comment));
      });
      sheet.getRange(rowNumber, 2, 1, row.length).setValues([row]);
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
  const headers = ["No", "Submit on", "Class", "House", "Form", "Cadet Name", "Cadet Number"];
  QUESTION_HEADERS.forEach(question => {
    headers.push(question, "Comment");
  });
  return headers;
}

function getNextResponseRow(sheet) {
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

function validateCadetName(value) {
  const name = validateOptionalText(value, "Cadet Name", 200);
  if (name && !/^[\p{L}\p{M}]+(?:\s+[\p{L}\p{M}]+)*$/u.test(name)) {
    throw new Error("Cadet Name may contain letters and spaces only.");
  }
  return name;
}

function validateCadetNumber(value) {
  const number = validateOptionalText(value, "Cadet Number", 100);
  if (number && !/^\d+$/.test(number)) {
    throw new Error("Cadet Number may contain digits only.");
  }
  return number;
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
      || !OPTIONS.includes(answer.selectedOption)
      || ((answer.selectedOption === "Satisfactory" || answer.selectedOption === "Poor")
        && !answer.comment.trim())) {
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
