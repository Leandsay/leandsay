/**
 * LEANDSAY LEE — JOB ENTRY WEB APP
 * Writes columns A:Q. Leaves O blank.
 * Preserves R:S calculated columns.
 */

const SPREADSHEET_ID = "1DasfUbog8nHiT4KuwkNbzQHJ_jsVAbE6M1Qlvq7ehQw";
const SHEET_GID = 1840782043;

const WORK_TYPES = [
  "JOB", "DOWNTIME", "PENDING", "CANCEL",
  "BREAK", "COACHING", "MEETING", "OUTAGE"
];

const JOB_TYPES = [
  "INITIAL", "REVISION", "UPLOAD", "DOWNTIME",
  "BREAK", "COACHING", "TRAINING"
];

function doGet() {
  return jsonResponse({ success: true, message: "Job Entry API is running." });
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    const data = getSubmittedData(e);
    validateEntry(data);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetById(SHEET_GID);
    if (!sheet) throw new Error("Target sheet tab not found. Check SHEET_GID.");

    const headerRow = findHeaderRow(sheet);
    if (!headerRow) throw new Error('Could not find "Designers" in column A.');

    // Find the last used row in A:Q only. R:S may contain formulas far below.
    const scanRows = Math.max(sheet.getLastRow(), headerRow);
    const values = sheet.getRange(headerRow + 1, 1, Math.max(1, scanRows - headerRow), 17)
      .getDisplayValues();

    let lastDataRow = headerRow;
    values.forEach((row, index) => {
      if (row.some(value => String(value).trim() !== "")) {
        lastDataRow = headerRow + 1 + index;
      }
    });

    const targetRow = lastDataRow + 1;

    // Copy formatting and validation from the previous data row.
    if (targetRow > headerRow + 1) {
      sheet.getRange(targetRow - 1, 1, 1, 19)
        .copyTo(sheet.getRange(targetRow, 1, 1, 19), { formatOnly: true });

      const rules = sheet.getRange(targetRow - 1, 1, 1, 19).getDataValidations();
      sheet.getRange(targetRow, 1, 1, 19).setDataValidations(rules);

      // Carry formulas in R:S down one row if formulas exist in the previous row.
      const formulaRange = sheet.getRange(targetRow - 1, 18, 1, 2);
      const formulas = formulaRange.getFormulasR1C1()[0];
      formulas.forEach((formula, i) => {
        if (formula) {
          sheet.getRange(targetRow, 18 + i).setFormulaR1C1(formula);
        }
      });
    }

    const row = [
      clean(data.designer),            // A Designers
      parseWorkDate(data.workDate),    // B Work Date
      clean(data.startTime),           // C Start
      clean(data.endTime),             // D End
      clean(data.jobId),               // E Job ID
      clean(data.customerName),        // F Customer Name
      clean(data.customerAddress),     // G Customer Address
      clean(data.state).toUpperCase(), // H State
      clean(data.mountingType),        // I M-Type
      clean(data.tool),                // J Tool
      clean(data.revision),            // K Rev #
      clean(data.categories),          // L Categories
      clean(data.salesRep),            // M Sales Rep
      clean(data.remarks),             // N Remarks
      "",                              // O Reserved blank
      clean(data.workType).toUpperCase(), // P Work Type
      clean(data.jobType).toUpperCase()   // Q Job Type
    ];

    sheet.getRange(targetRow, 1, 1, 17).setValues([row]);
    sheet.getRange(targetRow, 2).setNumberFormat("m/d/yyyy");
    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      message: "Entry saved.",
      row: targetRow
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.message || "Unable to save entry."
    });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function getSubmittedData(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  if (e && e.postData && e.postData.contents) {
    const type = e.postData.type || "";
    if (type.indexOf("application/json") !== -1) {
      return JSON.parse(e.postData.contents);
    }
  }
  throw new Error("No form payload received.");
}

function validateEntry(data) {
  ["designer", "workDate", "startTime", "endTime", "workType", "jobType"]
    .forEach(key => {
      if (!String(data[key] || "").trim()) {
        throw new Error(key + " is required.");
      }
    });

  if (![
    "COMBIS, MARK KEM M",
    "DIGAL, JADE C",
    "FLORES, JOHN MICHAEL S",
    "RECAMADAS, XYRA MAE E",
    "SIMILATAN, WILMA",
    "LENTIJAS, LEANDSAY LEE C"
  ].includes(data.designer)) {
    throw new Error("Please select a valid Designer.");
  }

  if (!WORK_TYPES.includes(String(data.workType).toUpperCase())) {
    throw new Error("Invalid Work Type.");
  }

  if (!JOB_TYPES.includes(String(data.jobType).toUpperCase())) {
    throw new Error("Invalid Job Type.");
  }

  if (!["ROOF", "GM", "BOTH"].includes(String(data.mountingType).toUpperCase())) {
    throw new Error("Invalid Mounting Type.");
  }

  if (!["SCANIFLY", "AURORA"].includes(String(data.tool).toUpperCase())) {
    throw new Error("Invalid Tool.");
  }

  if (String(data.workType).toUpperCase() === "JOB") {
    if (!String(data.jobId || "").trim()) throw new Error("Job ID is required for JOB.");
    if (!String(data.customerName || "").trim()) throw new Error("Customer Name is required for JOB.");
  }

  if (!validTime(data.startTime) || !validTime(data.endTime)) {
    throw new Error("Please enter valid start and end times.");
  }

  if (timeMinutes(data.endTime) < timeMinutes(data.startTime)) {
    throw new Error("End Time cannot be earlier than Start Time.");
  }
}

function validTime(value) {
  return /^([01]?\d|2[0-3]):[0-5]\d$/.test(String(value || "").trim()) ||
    /^(\d{1,2}):([0-5]\d)\s?(AM|PM)$/i.test(String(value || "").trim());
}

function timeMinutes(value) {
  const v = String(value).trim();
  let m = v.match(/^(\d{1,2}):(\d{2})\s?(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]) % 12;
    if (m[3].toUpperCase() === "PM") h += 12;
    return h * 60 + Number(m[2]);
  }
  m = v.match(/^(\d{1,2}):(\d{2})$/);
  return Number(m[1]) * 60 + Number(m[2]);
}

function parseWorkDate(value) {
  const s = String(value || "").trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error("Invalid Work Date.");
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function findHeaderRow(sheet) {
  const count = Math.min(10, sheet.getLastRow());
  if (count < 1) return null;
  const colA = sheet.getRange(1, 1, count, 1).getDisplayValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim().toLowerCase() === "designers") return i + 1;
  }
  return null;
}

function clean(value) {
  return String(value == null ? "" : value).trim();
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function testConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetById(SHEET_GID);
  if (!sheet) throw new Error("Target sheet tab not found.");
  Logger.log("Connected to: " + sheet.getName());
  Logger.log("Last row: " + sheet.getLastRow());
}
