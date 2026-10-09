/**
 * Menu item: run once at the end of a trimester. Exports the
 * working data sheets to CSV in the Archive Term ID folder, then
 * clears weekly values ONLY - advisor names, emails, divisions,
 * settings and the student roster are left in place so you don't
 * have to retype them if the same advisors/students continue.
 */
function archiveTrimester() {
  const ss = SpreadsheetApp.getActive();
  const archiveFolderId = getSetting('Archive Term ID (CSV archive folder)');
  if (!archiveFolderId) throw new Error('Set "Archive Term ID" in Settings first.');
  const folder = DriveApp.getFolderById(archiveFolderId);
  const term = getSetting('Term # (1, 2, or 3)');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  [SHEET_ROSTER, SHEET_ATTENDANCE, SHEET_MOM_LOG, SHEET_EMAIL_LOG].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const csv = sheetToCsv(sh);
    folder.createFile('T' + term + '_' + name.replace(/\s+/g, '_') + '_' + stamp + '.csv', csv, MimeType.CSV);
  });

  // Clear weekly values, keep structure/names
  clearWeeklyValues(ss.getSheetByName(SHEET_ATTENDANCE));
  clearWeeklyValues(ss.getSheetByName(SHEET_MOM_LOG));

  const roster = ss.getSheetByName(SHEET_ROSTER);
  if (roster) {
    const lastRow = roster.getLastRow();
    ['Week 3 Report Link', 'Week 6 Report Link', 'Week 9 Report Link'].forEach(label => {
      const header = roster.getRange(1, 1, 1, roster.getLastColumn()).getValues()[0];
      const col = header.indexOf(label);
      if (col !== -1 && lastRow > 1) roster.getRange(2, col + 1, lastRow - 1, 1).clearContent();
    });
  }

  SpreadsheetApp.getUi().alert('Archived Term ' + term + ' to CSV and cleared weekly values. ' +
    'Advisor names, emails and the student roster were kept. Update the Term # in Settings before starting the next trimester.');
}

function sheetToCsv(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.map(row => row.map(cell => {
    if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const s = String(cell).replace(/"/g, '""');
    return /[",\n]/.test(s) ? '"' + s + '"' : s;
  }).join(',')).join('\n');
}

function clearWeeklyValues(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 2) {
    sheet.getRange(2, 3, lastRow - 1, lastCol - 2).clearContent();
  } else if (lastRow > 1 && lastCol > 1) {
    sheet.getRange(2, 2, lastRow - 1, lastCol - 1).clearContent();
  }
}