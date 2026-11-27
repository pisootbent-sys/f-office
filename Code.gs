// ─────────────────────────────────────────
//  F Office Management — Code.gs  (เวอร์ชัน API สำหรับโฮสต์หน้าเว็บบน GitHub Pages)
//  Deploy เป็น Web app:  Execute as = Me  /  Who has access = Anyone
// ─────────────────────────────────────────

const SPREADSHEET_ID = '1l928kuU0kfriG5oQ1ScFnsbbkHMGh8tJqfX0JpDQ8k0';
const SHEET_NAME     = 'Sheet1';      // ตารางระบบ/ลิงก์ (คอลัมน์ A–J)
const SECTION_SHEET  = 'Sections';    // ตารางหัวข้อ (สร้างให้อัตโนมัติ)
const ADMIN_PIN      = '009';         // รหัสผู้ดูแล — ตรวจที่เซิร์ฟเวอร์ ผู้เข้าชมมองไม่เห็นรหัสนี้

// ── API entry points ─────────────────────
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  if (action === 'getAll') {
    return json_({ isOk: true, records: getRecords(), sections: getSections() });
  }
  return json_({ isOk: true, message: 'F Office API' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ isOk: false, error: 'bad request' });
  }

  // ทุกคำสั่งที่เขียนข้อมูลต้องมีรหัสผู้ดูแลที่ถูกต้อง
  if (String(req.pin || '') !== ADMIN_PIN) {
    return json_({ isOk: false, error: 'unauthorized' });
  }

  const p = req.payload;
  switch (req.action) {
    case 'verifyPin':     return json_({ isOk: true });
    case 'createRecord':  return json_(createRecord(p));
    case 'updateRecord':  return json_(updateRecord(p));
    case 'deleteRecord':  return json_(deleteRecord(p));
    case 'createSection': return json_(createSection(p));
    case 'updateSection': return json_(updateSection(p));
    case 'deleteSection': return json_(deleteSection(p));
    default:              return json_({ isOk: false, error: 'unknown action' });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
    section:     row[7] ? String(row[7]) : '',
    extraLabel:  row[8] ? String(row[8]) : '',
    extraUrl:    row[9] ? String(row[9]) : ''
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

// ลบหัวข้อ: ระบบในหัวข้อนี้ไม่ถูกลบ แต่จะย้ายไป "ระบบอื่นๆ"
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
