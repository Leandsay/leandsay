/** Aurora Job Entry Form — Google Apps Script backend.
 * Attach this project to the Google Sheet that receives entries.
 * Columns A:P match the supplied log: Designers, Work Date, Start, End,
 * Job ID, Customer Name, Customer Address, State, M-Type, Tool, Rev #,
 * Categories, Sales Rep, Remarks, reserved blank column O, Work Type.
 */
const CONFIG = {
  SHEET_NAME: '', // Optional exact tab name. Blank = find tab by headers.
  HEADER_SEARCH_ROWS: 10,
  DATE_FORMAT: 'm/d/yyyy',
  TIME_FORMAT: 'h:mm:ss AM/PM'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Aurora Job Entry Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = findTargetSheet_(ss);
  const optionCols = { designers: 1, states: 8, mTypes: 9, tools: 10, revisions: 11, categories: 12, salesReps: 13 };
  const options = {};
  Object.keys(optionCols).forEach(k => options[k] = getOptions_(target.sheet, target.headerRow, optionCols[k]));
  return {
    options,
    today: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'Asia/Manila', 'yyyy-MM-dd'),
    sheetName: target.sheet.getName()
  };
}

function submitJobEntry(data) {
  if (!data) throw new Error('No form data received.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const target = findTargetSheet_(ss);
    const sheet = target.sheet;
    const designer = clean_(data.designer);
    const workDate = clean_(data.workDate);
    const startTime = clean_(data.startTime);
    const endTime = clean_(data.endTime);
    const workType = clean_(data.workType).toUpperCase();
    if (!designer || !workDate || !startTime || !endTime || !workType) {
      throw new Error('Complete Designer, Work Date, Start Time, End Time, and Work Type.');
    }
    if (!['JOB', 'BREAK', 'DOWNTIME'].includes(workType)) throw new Error('Invalid Work Type.');
    if (workType === 'JOB' && (!clean_(data.jobId) || !clean_(data.customerName))) {
      throw new Error('Job ID and Customer Name are required for JOB entries.');
    }
    const date = parseDate_(workDate);
    const start = combineDateTime_(date, startTime);
    const end = combineDateTime_(date, endTime);
    if (end.getTime() < start.getTime()) throw new Error('End Time cannot be earlier than Start Time. For overnight work, use the next date.');

    // A:P. O is deliberately left blank to preserve the existing sheet layout.
    const row = [
      designer, date, start, end,
      clean_(data.jobId), clean_(data.customerName), clean_(data.customerAddress),
      clean_(data.state), clean_(data.mType), clean_(data.tool), clean_(data.revision),
      clean_(data.category), clean_(data.salesRep), clean_(data.remarks), '', workType
    ];
    const nextRow = Math.max(sheet.getLastRow() + 1, target.headerRow + 1);
    sheet.getRange(nextRow, 1, 1, 16).setValues([row]);
    sheet.getRange(nextRow, 2).setNumberFormat(CONFIG.DATE_FORMAT);
    sheet.getRange(nextRow, 3, 1, 2).setNumberFormat(CONFIG.TIME_FORMAT);
    return { success: true, rowNumber: nextRow, sheetName: sheet.getName() };
  } finally {
    lock.releaseLock();
  }
}

function findTargetSheet_(ss) {
  if (CONFIG.SHEET_NAME) {
    const s = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!s) throw new Error('Sheet tab not found: ' + CONFIG.SHEET_NAME);
    const h = findHeaderRow_(s);
    if (!h) throw new Error('Expected Aurora headers were not found in ' + CONFIG.SHEET_NAME);
    return { sheet: s, headerRow: h };
  }
  for (const s of ss.getSheets()) {
    const h = findHeaderRow_(s);
    if (h) return { sheet: s, headerRow: h };
  }
  throw new Error('Could not locate the Aurora log tab. Check that the header row matches the supplied spreadsheet layout.');
}

function findHeaderRow_(sheet) {
  const rows = Math.min(CONFIG.HEADER_SEARCH_ROWS, sheet.getMaxRows());
  const values = sheet.getRange(1, 1, rows, 16).getDisplayValues();
  const expected = {
    0: ['designers', 'designer'], 2: ['start'], 3: ['end'], 4: ['job id', 'jobid'],
    5: ['customer name'], 6: ['customer address'], 7: ['state'],
    8: ['m-type', 'm type', 'mtype'], 9: ['tool'], 10: ['rev #', 'rev', 'revision'],
    11: ['categories', 'category'], 12: ['sales rep', 'sales representative'],
    13: ['remarks', 'remark'], 15: ['work type', 'worktype']
  };
  for (let r = 0; r < values.length; r++) {
    let matches = 0;
    Object.keys(expected).forEach(i => {
      const v = normalize_(values[r][Number(i)]);
      if (expected[i].includes(v)) matches++;
    });
    if (matches >= 12) return r + 1;
  }
  return null;
}

function getOptions_(sheet, headerRow, col) {
  const cell = sheet.getRange(headerRow + 1, col);
  const rule = cell.getDataValidation();
  if (rule) {
    const type = rule.getCriteriaType();
    const args = rule.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST && args[0]) return unique_(args[0]);
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && args[0]) {
      const list = unique_(args[0].getDisplayValues().flat());
      if (list.length) return list;
    }
  }
  const last = sheet.getLastRow();
  if (last <= headerRow) return [];
  return unique_(sheet.getRange(headerRow + 1, col, last - headerRow, 1).getDisplayValues().flat());
}

function parseDate_(value) {
  const p = value.split('-').map(Number);
  if (p.length !== 3 || p.some(n => !Number.isFinite(n))) throw new Error('Invalid work date.');
  return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
}
function combineDateTime_(date, value) {
  const p = value.split(':').map(Number);
  if (p.length < 2 || p.some(n => !Number.isFinite(n)) || p[0] < 0 || p[0] > 23 || p[1] < 0 || p[1] > 59) throw new Error('Invalid time.');
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), p[0], p[1], 0);
}
function clean_(v) { return String(v == null ? '' : v).trim(); }
function normalize_(v) { return clean_(v).toLowerCase().replace(/\s+/g, ' '); }
function unique_(values) {
  const seen = new Set();
  return values.map(clean_).filter(v => v && !seen.has(v) && seen.add(v));
}
