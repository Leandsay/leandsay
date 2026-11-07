
/**
 * LEANDSAY LEE — JOB ENTRY WEB APP
 * Writes columns A:Q. Leaves O blank.
 * Preserves R:S calculated columns.
 */

const SPREADSHEET_ID = "1DasfUbog8nHiT4KuwkNbzQHJ_jsVAbE6M1Qlvq7ehQw";
const SHEET_GID = 885676598;

const WORK_TYPES = [
  "JOB", "DOWNTIME", "PENDING", "CANCEL",
  "BREAK", "COACHING", "MEETING", "OUTAGE"
];

const JOB_TYPES = [
  "INITIAL", "REVISION", "UPLOAD", "DOWNTIME",
  "BREAK", "COACHING", "TRAINING"
];

// Exact names sent by the HTML form, mapped to spreadsheet names.
const DESIGNER_MAP = {
  "COMBIS, MARK KEM M": "COMBIS MARK KEM",
  "DIGAL, JADE C": "DIGAL JADE",
  "FLORES, JOHN MICHAEL S": "FLORES JOHN MICHAEL",
  "RECAMADAS, XYRA MAE E": "RECAMADAS XYRA MAE",
  "SIMILATAN, WILMA": "SIMILATAN WILMA",
  "LENTIJAS, LEANDSAY LEE C": "LENTIJAS LEANDSAY LEE"
};

function doGet() {
  return jsonResponse({
    success: true,
    message: "Job Entry API is running."
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    const data = getSubmittedData(e);
    validateEntry(data);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetById(SHEET_GID);

    if (!sheet) {
      throw new Error("Target sheet tab not found. Please check SHEET_GID.");
    }

    const headerRow = findHeaderRow(sheet);

    if (!headerRow) {
      throw new Error('Could not find "Designers" in column A.');
    }

    // Find the last used row in columns A:Q.
    const scanRows = Math.max(sheet.getLastRow(), headerRow);
    const numberOfRows = Math.max(1, scanRows - headerRow);

    const values = sheet
      .getRange(headerRow + 1, 1, numberOfRows, 17)
      .getDisplayValues();

    let lastDataRow = headerRow;

    values.forEach((row, index) => {
      if (row.some(value => String(value).trim() !== "")) {
        lastDataRow = headerRow + 1 + index;
      }
    });

    const targetRow = lastDataRow + 1;

    // Preserve formatting and dropdown validations.
    if (targetRow > headerRow + 1) {
      const previousRow = sheet.getRange(targetRow - 1, 1, 1, 19);
      const newRow = sheet.getRange(targetRow, 1, 1, 19);

      previousRow.copyTo(newRow, { formatOnly: true });
      newRow.setDataValidations(previousRow.getDataValidations());

      // Carry formulas in R:S down to the new row.
      const formulas = sheet
        .getRange(targetRow - 1, 18, 1, 2)
        .getFormulasR1C1()[0];

      formulas.forEach((formula, index) => {
        if (formula) {
          sheet.getRange(targetRow, 18 + index)
            .setFormulaR1C1(formula);
        }
      });
    }

    const designerKey = clean(data.designer).toUpperCase();
    const spreadsheetDesigner = DESIGNER_MAP[designerKey];

    const row = [
      spreadsheetDesigner,                  // A Designer
      parseWorkDate(data.workDate),          // B Work Date
      timeToSheetValue(data.startTime),      // C Start
      timeToSheetValue(data.endTime),        // D End
      clean(data.jobId),                     // E Job ID
      clean(data.customerName),              // F Customer Name
      clean(data.customerAddress),           // G Customer Address
      clean(data.state).toUpperCase(),       // H State
      clean(data.mountingType).toUpperCase(),// I M-Type
      clean(data.tool).toUpperCase(),        // J Tool
      clean(data.revision),                  // K Rev #
      clean(data.categories),                // L Categories
      clean(data.salesRep),                  // M Sales Rep
      clean(data.remarks),                   // N Remarks
      "",                                    // O Reserved blank
      clean(data.workType).toUpperCase(),    // P Work Type
      clean(data.jobType).toUpperCase()      // Q Job Type
    ];

    sheet.getRange(targetRow, 1, 1, 17).setValues([row]);

    // Apply date and time formats.
    sheet.getRange(targetRow, 2).setNumberFormat("m/d/yyyy");
    sheet.getRange(targetRow, 3, 1, 2)
      .setNumberFormat("h:mm AM/PM");

    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      message: "Entry saved successfully.",
      row: targetRow
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.message || "Unable to save entry."
    });

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function getSubmittedData(e) {
  if (e && e.parameter && e.parameter.payload) {
    try {
      return JSON.parse(e.parameter.payload);
    } catch (error) {
      throw new Error("Invalid form payload. Please check the submitted data.");
    }
  }

  if (e && e.postData && e.postData.contents) {
    const type = e.postData.type || "";

    if (type.indexOf("application/json") !== -1) {
      try {
        return JSON.parse(e.postData.contents);
      } catch (error) {
        throw new Error("Invalid JSON submission.");
      }
    }
  }

  throw new Error("No form payload received.");
}

function validateEntry(data) {
  if (!data || typeof data !== "object") {
    throw new Error("Invalid submission data.");
  }

  [
    "designer",
    "workDate",
    "startTime",
    "endTime",
    "workType",
    "jobType"
  ].forEach(key => {
    if (!clean(data[key])) {
      throw new Error(key + " is required.");
    }
  });

  const designerKey = clean(data.designer).toUpperCase();

  if (!DESIGNER_MAP[designerKey]) {
    throw new Error("Please select a valid Designer.");
  }

  const workType = clean(data.workType).toUpperCase();
  const jobType = clean(data.jobType).toUpperCase();

  if (!WORK_TYPES.includes(workType)) {
    throw new Error("Invalid Work Type.");
  }

  if (!JOB_TYPES.includes(jobType)) {
    throw new Error("Invalid Job Type.");
  }

  // Mounting Type and Tool are required for JOB entries.
  if (workType === "JOB") {
    if (!["ROOF", "GM", "BOTH"].includes(
      clean(data.mountingType).toUpperCase()
    )) {
      throw new Error("Please select a valid Mounting Type.");
    }

    if (!["SCANIFLY", "AURORA"].includes(
      clean(data.tool).toUpperCase()
    )) {
      throw new Error("Please select a valid Tool.");
    }

    if (!clean(data.jobId)) {
      throw new Error("Job ID is required for JOB.");
    }

    if (!clean(data.customerName)) {
      throw new Error("Customer Name is required for JOB.");
    }
  } else {
    // If supplied for non-JOB entries, values must still be valid.
    const mountingType = clean(data.mountingType).toUpperCase();
    const tool = clean(data.tool).toUpperCase();

    if (mountingType && !["ROOF", "GM", "BOTH"].includes(mountingType)) {
      throw new Error("Invalid Mounting Type.");
    }

    if (tool && !["SCANIFLY", "AURORA"].includes(tool)) {
      throw new Error("Invalid Tool.");
    }
  }

  if (!validTime(data.startTime) || !validTime(data.endTime)) {
    throw new Error("Please enter valid start and end times.");
  }

  if (timeMinutes(data.endTime) < timeMinutes(data.startTime)) {
    throw new Error("End Time cannot be earlier than Start Time.");
  }

  // Confirm Work Date uses YYYY-MM-DD format.
  if (!/^\d{4}-\d{2}-\d{2}$/.test(clean(data.workDate))) {
    throw new Error("Invalid Work Date. Expected YYYY-MM-DD.");
  }
}

function validTime(value) {
  const v = clean(value);

  return /^([01]?\d|2[0-3]):[0-5]\d$/.test(v) ||
    /^(\d{1,2}):([0-5]\d)\s?(AM|PM)$/i.test(v);
}

function timeToSheetValue(value) {
  return timeMinutes(value) / 1440;
}

function timeMinutes(value) {
  const v = clean(value);

  let match = v.match(/^(\d{1,2}):(\d{2})\s?(AM|PM)$/i);

  if (match) {
    let hour = Number(match[1]) % 12;

    if (match[3].toUpperCase() === "PM") {
      hour += 12;
    }

    return hour * 60 + Number(match[2]);
  }

  match = v.match(/^(\d{1,2}):(\d{2})$/);

  if (!match) {
    throw new Error("Invalid time format.");
  }

  return Number(match[1]) * 60 + Number(match[2]);
}

function parseWorkDate(value) {
  const match = clean(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    throw new Error("Invalid Work Date.");
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  const date = new Date(year, month - 1, day, 12, 0, 0);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error("Invalid calendar date.");
  }

  return date;
}

function findHeaderRow(sheet) {
  const count = Math.min(10, sheet.getLastRow());

  if (count < 1) {
    return null;
  }

  const columnA = sheet.getRange(1, 1, count, 1).getDisplayValues();

  for (let i = 0; i < columnA.length; i++) {
    if (clean(columnA[i][0]).toLowerCase() === "designers") {
      return i + 1;
    }
  }

  return null;
}

function clean(value) {
  return String(value == null ? "" : value).trim();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function testConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetById(SHEET_GID);

  if (!sheet) {
    throw new Error("Target sheet tab not found. Check SHEET_GID.");
  }

  Logger.log("Connected to: " + sheet.getName());
  Logger.log("Last row: " + sheet.getLastRow());
}
