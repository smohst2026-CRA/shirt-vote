/**
 * Google Apps Script backend สำหรับเว็บโหวตแบบเสื้อ
 * วิธี deploy ดูในไฟล์ "วิธีตั้งค่า.md"
 */

const SHEET_NAME = "Votes";
const DESIGN_COUNT = 13; // จำนวนแบบเสื้อ

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // ตั้งค่าเริ่มต้นถ้ายังไม่มีข้อมูล
  if (sheet.getRange("A1").getValue() === "") {
    sheet.getRange("A1:B1").setValues([["design_id", "votes"]]);
    const rows = [];
    for (let i = 1; i <= DESIGN_COUNT; i++) {
      rows.push([i, 0]);
    }
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  // เติมแถวให้กับแบบเสื้อที่เพิ่มเข้ามาใหม่ (ถ้ายังไม่มี)
  const lastRow = sheet.getLastRow();
  const existingIds = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => Number(r[0]))
    : [];
  const newRows = [];
  for (let i = 1; i <= DESIGN_COUNT; i++) {
    if (existingIds.indexOf(i) === -1) newRows.push([i, 0]);
  }
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, 2).setValues(newRows);
  }

  return sheet;
}

function getVotes_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const votes = {};
  data.forEach(row => {
    votes[row[0]] = row[1];
  });
  return votes;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "results") {
    return jsonOutput_({ votes: getVotes_() });
  }
  return jsonOutput_({ error: "unknown action" });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet_();

  if (body.action === "vote") {
    const id = Number(body.id);
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (Number(ids[i][0]) === id) {
        const cell = sheet.getRange(i + 2, 2);
        cell.setValue(cell.getValue() + 1);
        break;
      }
    }
    return jsonOutput_({ votes: getVotes_() });
  }

  if (body.action === "reset") {
    const lastRow = sheet.getLastRow();
    const zeros = [];
    for (let i = 0; i < lastRow - 1; i++) zeros.push([0]);
    sheet.getRange(2, 2, zeros.length, 1).setValues(zeros);
    return jsonOutput_({ votes: getVotes_() });
  }

  return jsonOutput_({ error: "unknown action" });
}
