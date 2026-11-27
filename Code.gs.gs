// ─────────────────────────────────────────
//  F Office Management — Code.gs
// ─────────────────────────────────────────

const SPREADSHEET_ID = '1l928kuU0kfriG5oQ1ScFnsbbkHMGh8tJqfX0JpDQ8k0';
const SHEET_NAME     = 'Sheet1';      // ตารางระบบ/ลิงก์ (คอลัมน์ A–J)
const SECTION_SHEET  = 'Sections';    // ตารางหัวข้อ (สร้างให้อัตโนมัติ)

// ── Entry Point ──────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('F Office Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Read ─────────────────────────────────
function getRecords() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  data.shift(); // remove header row

  return data.map(row => ({
    __backendId: String(row[0]),
    id:          String(row[0]),
    title:       row[1],
    url:         row[2],
    icon:        row[3],
    color:       row[4],
    order:       row[5],
    createdAt:   row[6],
    section:     row[7] ? String(row[7]) : '',  // รหัสหัวข้อ (คอลัมน์ H)
    extraLabel:  row[8] ? String(row[8]) : '',  // ชื่อลิงก์ภายนอก (คอลัมน์ I)
    extraUrl:    row[9] ? String(row[9]) : ''   // URL ลิงก์ภายนอก (คอลัมน์ J)
  }));
}

// ── Create ───────────────────────────────
function createRecord(record) {
  try {
    getSheet().appendRow([
      record.id,
      record.title,
      record.url,
      record.icon,
      record.color,
      record.order,
      record.createdAt,
      record.section || '',
      record.extraLabel || '',
      record.extraUrl || ''
    ]);
    return { isOk: true };
  } catch (e) {
    return { isOk: false, error: e.toString() };
  }
}

// ── Update ───────────────────────────────
function updateRecord(record) {
  try {
    const sheet    = getSheet();
    const data     = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]) === String(record.id));

    if (rowIndex === -1) return { isOk: false, error: 'Record not found' };

    sheet.getRange(rowIndex + 1, 1, 1, 10).setValues([[
      record.id,
      record.title,
      record.url,
      record.icon,
      record.color,
      record.order,
      record.createdAt,
      record.section || '',
      record.extraLabel || '',
      record.extraUrl || ''
    ]]);
    return { isOk: true };
  } catch (e) {
    return { isOk: false, error: e.toString() };
  }
}

// ── Delete ───────────────────────────────
function deleteRecord(record) {
  try {
    const sheet    = getSheet();
    const data     = sheet.getDataRange().getValues();
    const targetId = String(record.__backendId || record.id);
    const rowIndex = data.findIndex(row => String(row[0]) === targetId);

    if (rowIndex === -1) return { isOk: false, error: 'Record not found' };

    sheet.deleteRow(rowIndex + 1);
    return { isOk: true };
  } catch (e) {
    return { isOk: false, error: e.toString() };
  }
}

// ═════════════════════════════════════════
//  หัวข้อ (Sections)
// ═════════════════════════════════════════
function getSections() {
  const data = getSectionSheet().getDataRange().getValues();
  data.shift(); // remove header row

  return data
    .filter(row => row[0] !== '')
    .map(row => ({
      __backendId: String(row[0]),
      id:          String(row[0]),
      name:        row[1],
      order:       row[2],
      createdAt:   row[3]
    }));
}

function createSection(section) {
  try {
    getSectionSheet().appendRow([
      section.id,
      section.name,
      section.order,
      section.createdAt
    ]);
    return { isOk: true };
  } catch (e) {
    return { isOk: false, error: e.toString() };
  }
}

function updateSection(section) {
  try {
    const sheet    = getSectionSheet();
    const data     = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]) === String(section.id));

    if (rowIndex === -1) return { isOk: false, error: 'Section not found' };

    sheet.getRange(rowIndex + 1, 1, 1, 4).setValues([[
      section.id,
      section.name,
      section.order,
      section.createdAt
    ]]);
    return { isOk: true };
  } catch (e) {
    return { isOk: false, error: e.toString() };
  }
}

// ลบหัวข้อ: ระบบที่อยู่ในหัวข้อนี้ไม่ถูกลบ แต่จะย้ายไป "ระบบอื่นๆ"
function deleteSection(section) {
  try {
    const targetId = String(section.__backendId || section.id);

    const secSheet = getSectionSheet();
    const secData  = secSheet.getDataRange().getValues();
    const secIndex = secData.findIndex(row => String(row[0]) === targetId);
    if (secIndex === -1) return { isOk: false, error: 'Section not found' };
    secSheet.deleteRow(secIndex + 1);

    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][7]) === targetId) {
        sheet.getRange(i + 1, 8).setValue('');
      }
    }
    return { isOk: true };
  } catch (e) {
    return { isOk: false, error: e.toString() };
  }
}

// ── Helpers ──────────────────────────────
function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function getSectionSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SECTION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SECTION_SHEET);
    sheet.appendRow(['id', 'name', 'order', 'createdAt']);
  }
  return sheet;
}
