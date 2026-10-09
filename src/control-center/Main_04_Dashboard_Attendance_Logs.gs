/**
 * Menu item: reads the current week's student rows out of each
 * advisor's own workbook (if reachable) and writes a same-page
 * summary here, so you can see where everyone stands without
 * opening six separate files. Manual, on-demand - not automatic,
 * since Sheets can't push changes from one file into another live.
 */
function pullLiveStatusFromAdvisors() {
  const ss = SpreadsheetApp.getActive();
  const tw = getCurrentTermWeek();
  if (!tw) {
    SpreadsheetApp.getUi().alert('No active advisory week right now (break/exam period) - nothing to pull.');
    return;
  }
  let sh = ss.getSheetByName('Live Status');
  if (!sh) sh = ss.insertSheet('Live Status');
  sh.clear();
  safeMerge_(sh, 'A1:F1').setValue('Live Status - Term ' + tw.term + ', Week ' + tw.week +
    ' (pulled ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM, HH:mm') + ')')
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  const header = ['Advisor', 'Student', 'Weekly Status', 'On Track?', 'Flag Counsellor?', 'Notes'];
  sh.getRange(2, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_TABLE_HEAD);

  const folderId = getSetting('Folder Term ID (advisor workbooks live here)');
  let row = 3;
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    getAdvisorList().forEach(a => {
      const data = getStudentDataForWeek(a.name, tw.week);
      data.forEach(s => {
        sh.getRange(row, 1, 1, 6).setValues([[a.name, s.name, s.status, '', '', s.concerns]]);
        row++;
      });
    });
  }
  addStatusColorRule(sh, sh.getRange(3, 3, Math.max(row - 3, 1), 1), {
    'Better': STATUS_COMPLETED, 'Same': STATUS_IN_PROGRESS, 'Worse': STATUS_OVERDUE,
  });
  sh.getDataRange().setWrap(true);
  sh.setFrozenRows(2);
  SpreadsheetApp.getUi().alert('Live Status tab refreshed from ' + (row - 3) + ' student rows across all advisors.');
}

// ---- DASHBOARD ----
function setupDashboardSheet(ss) {
  let sh = ss.getSheetByName(SHEET_DASHBOARD);
  if (!sh) sh = ss.insertSheet(SHEET_DASHBOARD, 0);
  buildDashboardLayout(sh);
  refreshDashboard();
}

/** Builds the static parts of the dashboard once: title band, KPI
 *  card shells, section headers, and quick-jump links. refreshDashboard()
 *  fills in the numbers/rows every time it's run. */
function buildDashboardLayout(sh) {
  sh.clear();
  sh.clearConditionalFormatRules();

  // ---- Title band ----
  safeMerge_(sh, 'A1:H1').setValue('🌱 Advisor Dashboard')
    .setFontSize(20).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  sh.setRowHeight(1, 40);
  safeMerge_(sh, 'A2:H2').setValue('Your at-a-glance view - refresh any time from the menu.')
    .setFontStyle('italic').setFontColor(COLOR_READONLY_TXT);

  // ---- KPI cards (row 4-5), 6 across ----
  const kpiLabels = ['Today', 'Current Week', 'Advisors', 'Students', 'Meetings Pending', 'Flagged This Week'];
  for (let i = 0; i < kpiLabels.length; i++) {
    const col = i + 1;
    sh.getRange(4, col).setValue(kpiLabels[i]).setFontWeight('bold')
      .setFontColor(COLOR_KPI_TXT).setBackground(COLOR_KPI_BG).setHorizontalAlignment('center').setWrap(true);
    sh.getRange(5, col).setFontSize(14).setFontWeight('bold').setFontColor(COLOR_KPI_TXT)
      .setBackground(COLOR_KPI_BG).setHorizontalAlignment('center');
  }
  sh.setRowHeight(4, 34);
  sh.setRowHeight(5, 34);

  // ---- Note banner ----
  safeMerge_(sh, 'A7:H7').setBackground(COLOR_SECTION_BG).setFontColor(COLOR_SECTION_TXT)
    .setFontWeight('bold').setWrap(true);

  // ---- Quick links section ----
  sh.getRange('A9').setValue('Quick Links').setFontWeight('bold').setFontSize(12);
  refreshQuickLinks(sh);

  // ---- Upcoming meetings table ----
  sh.getRange('A12').setValue('📅 Meetings Not Yet Marked Complete').setFontWeight('bold').setFontSize(12);
  sh.getRange(13, 1, 1, 3).setValues([['Advisor', 'Week', 'Status']])
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_TABLE_HEAD);

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(2);
}

function refreshQuickLinks(sh) {
  const linkLabels = ['Advisor Roster', 'Advisor Attendance', 'MoM Log', 'Live Status', 'Trimester Reference', 'Email Log'];
  linkLabels.forEach((name, i) => {
    const target = SpreadsheetApp.getActive().getSheetByName(name);
    const cell = sh.getRange(10, i + 1);
    if (target) {
      cell.setFormula('=HYPERLINK("#gid=' + target.getSheetId() + '","' + name + '")');
    } else {
      cell.setValue(name + ' (not built yet)').setFontColor(COLOR_READONLY_TXT);
    }
    cell.setBackground(COLOR_KPI_BG).setHorizontalAlignment('center').setWrap(true);
  });
  sh.setRowHeight(10, 28);
}

/** Menu item: recomputes today's date, week number, KPI numbers, and
 *  the upcoming-meetings list. Run any time - safe, non-destructive. */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_DASHBOARD);
  if (!sh) return;
  refreshQuickLinks(sh);

  const tw = getCurrentTermWeek();
  const advisorCount = getAdvisorList().length;
  const studentCount = getStudentList().length;

  // Meetings pending = advisors whose current week isn't Completed
  const att = ss.getSheetByName(SHEET_ATTENDANCE);
  let pending = 0;
  const upcoming = [];
  if (att && tw) {
    const data = att.getDataRange().getValues();
    const header = data[ROSTER_HEADER_ROW - 1];
    const weekCol = header.indexOf('Week ' + tw.week);
    if (weekCol !== -1) {
      for (let r = ROSTER_HEADER_ROW; r < data.length; r++) {
        const status = data[r][weekCol] || 'Scheduled';
        if (data[r][0] && status !== 'Completed') {
          pending++;
          upcoming.push([data[r][0], 'Week ' + tw.week, status]);
        }
      }
    }
  }

  // Flagged this week = students marked "Not on Track" or Flag Counsellor,
  // pulled live from advisor workbooks if reachable (same as Live Status)
  let flagged = '—';
  const folderId = getSetting('Folder Term ID (advisor workbooks live here)');
  if (folderId && tw) {
    let count = 0;
    getAdvisorList().forEach(a => {
      const data = getStudentDataForWeek(a.name, tw.week);
      data.forEach(s => { if (s.status && s.status.indexOf('Worse') !== -1) count++; });
    });
    flagged = String(count);
  }

  const kpiValues = [
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE, d MMM'),
    tw ? 'T' + tw.term + ' - W' + tw.week : 'Break/Exam',
    String(advisorCount),
    String(studentCount),
    String(pending),
    flagged,
  ];
  sh.getRange(5, 1, 1, kpiValues.length).setValues([kpiValues]);

  sh.getRange('A7').setValue(
    tw
      ? (tw.note ? '📌 ' + tw.note : '✅ On track - no school-calendar exceptions this week.')
      : '⏸ No active advisory week right now (break or exam period) - reschedule any pending meetings online.'
  );

  const startRow = 14;
  const clearRows = Math.max(sh.getLastRow() - startRow + 1, 1);
  sh.getRange(startRow, 1, clearRows, 3).clearContent();
  sh.clearConditionalFormatRules();
  if (upcoming.length) {
    sh.getRange(startRow, 1, upcoming.length, 3).setValues(upcoming);
    for (let i = 0; i < upcoming.length; i++) {
      if (i % 2 === 1) sh.getRange(startRow + i, 1, 1, 3).setBackground(COLOR_ALT_ROW);
    }
    addStatusColorRule(sh, sh.getRange(startRow, 3, upcoming.length, 1), {
      'Scheduled': STATUS_PENDING, 'Physical': FUNC_INFO, 'Virtual': FUNC_INFO,
      'Day Off - Rescheduled Online': STATUS_AT_RISK,
    });
  } else {
    sh.getRange(startRow, 1).setValue('🎉 All meetings marked complete for this week.').setFontColor(COLOR_KPI_TXT);
  }
}

/**
 * Installable onEdit trigger for the MAIN workbook (separate from the
 * one in each advisor's own file). Watches the Attendance tab: the
 * moment you mark a week "Completed" for an advisor, "Attend Weekly
 * Meeting" flips to Done on both their profile tab here and their
 * own Week N tab - no need to send the MoM first for that to happen.
 */
function enableMainLiveSync() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'mainWorkbookOnEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mainWorkbookOnEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert(
    'Live sync enabled. Marking a week "Completed" in Attendance now instantly flips ' +
    '"Attend Weekly Meeting" to Done on that advisor\'s profile - in this workbook and their own.'
  );
}

function mainWorkbookOnEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_ATTENDANCE) return;
    if (e.range.getRow() < ROSTER_DATA_START_ROW || e.range.getColumn() < 3) return;
    if (e.value !== 'Completed') return;

    const advisorName = sh.getRange(e.range.getRow(), 1).getValue();
    const header = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0];
    const weekLabel = header[e.range.getColumn() - 1]; // e.g. "Week 3"
    const match = String(weekLabel).match(/Week (\d+)/);
    if (!advisorName || !match) return;

    pushAttendDone(advisorName, parseInt(match[1], 10));
  } catch (err) {
    // never let a trigger error block editing
  }
}

// ---- ATTENDANCE (AD status) ----
/** Rebuilds the Attendance tab from a given advisor name list (pulled
 *  from the current Roster, not a hardcoded list) - so it always
 *  reflects whoever's actually on the roster this trimester. Any
 *  week-status values already entered for an advisor who's still on
 *  the roster are preserved; advisors no longer on the roster are
 *  dropped, new ones are added with a fresh "Scheduled" row. */
function rebuildAttendanceSheet(ss, advisorRows) {
  let sh = ss.getSheetByName(SHEET_ATTENDANCE);
  const preserved = sh ? captureWeeklyGrid_(sh) : {};
  if (!sh) sh = ss.insertSheet(SHEET_ATTENDANCE);
  sh.clear();
  sh.clearConditionalFormatRules();
  clearAllValidations_(sh);

  bandTitle_(sh, 'A1:L1', 'Advisor Attendance');
  sh.setRowHeight(1, 32);

  const weekCount = getWeekCount();
  const header = ['Advisor Name', 'Email'];
  for (let w = 1; w <= weekCount; w++) header.push('Week ' + w);
  sh.getRange(2, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground(COLOR_TABLE_HEAD).setFontColor(COLOR_HEADER_TEXT);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Scheduled', 'Day Off - Rescheduled Online', 'Completed'], true)
    .setAllowInvalid(false)
    .build();

  advisorRows.forEach((a, i) => {
    const r = i + 3;
    const prevRow = preserved[a.name]; // index0 = Email, index1.. = week values
    sh.getRange(r, 1).setValue(a.name);
    sh.getRange(r, 2).setValue(a.email || (prevRow ? prevRow[0] : '') || '');
    const rowVals = [];
    for (let w = 1; w <= weekCount; w++) rowVals.push(prevRow ? (prevRow[w] || 'Scheduled') : 'Scheduled');
    sh.getRange(r, 3, 1, weekCount).setValues([rowVals]).setDataValidation(rule);
  });

  if (advisorRows.length) {
    addStatusColorRule(sh, sh.getRange(3, 3, advisorRows.length, weekCount), {
      'Scheduled': STATUS_PENDING,
      'Day Off - Rescheduled Online': STATUS_AT_RISK,
      'Completed': STATUS_COMPLETED,
    });
    try { sh.getRange(3, 1, advisorRows.length, header.length).applyRowBanding(SpreadsheetApp.BandingTheme.TEAL, false, false); } catch (e) {}
  }

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
}

/** Reads an existing "Advisor Name | Week 1..N" grid (header on row 2,
 *  data from row 3) into {advisorName: [week1val, ...]} before a rebuild. */
function captureWeeklyGrid_(sheet) {
  const out = {};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < ROSTER_DATA_START_ROW || lastCol < 2) return out;
  const data = sheet.getRange(ROSTER_DATA_START_ROW, 1, lastRow - ROSTER_DATA_START_ROW + 1, lastCol).getValues();
  data.forEach(row => {
    if (row[0]) out[row[0]] = row.slice(1);
  });
  return out;
}

// ---- MoM LOG (also feeds Email Log + child-workbook mirror) ----
/** Rebuilds the MoM Log tab from the current Roster's advisor rows
 *  (name + email), preserving any MoM links already logged for
 *  advisors who are still on the roster. */
function rebuildMoMLogSheet(ss, advisorRows) {
  let sh = ss.getSheetByName(SHEET_MOM_LOG);
  const preserved = sh ? captureWeeklyGrid_(sh) : {}; // keyed by name, values start at Email then Week1..N
  if (!sh) sh = ss.insertSheet(SHEET_MOM_LOG);
  sh.clear();
  sh.clearConditionalFormatRules();
  clearAllValidations_(sh);

  bandTitle_(sh, 'A1:L1', 'MoM Log');
  sh.setRowHeight(1, 32);

  const weekCount = getWeekCount();
  const header = ['Advisor Name', 'Email'];
  for (let w = 1; w <= weekCount; w++) header.push('Week ' + w + ' MoM Link');
  sh.getRange(2, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground(COLOR_TABLE_HEAD).setFontColor(COLOR_HEADER_TEXT);

  advisorRows.forEach((a, i) => {
    const r = i + 3;
    const prevRow = preserved[a.name]; // index0 = Email, index1.. = week links
    const email = a.email || (prevRow ? prevRow[0] : '') || '';
    const links = [];
    for (let w = 1; w <= weekCount; w++) links.push(prevRow ? (prevRow[w] || '') : '');
    sh.getRange(r, 1, 1, header.length).setValues([[a.name, email].concat(links)]);
  });

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
}

// ---- EMAIL LOG ----
function setupEmailLogSheet(ss) {
  let sh = ss.getSheetByName(SHEET_EMAIL_LOG);
  if (sh) return;
  sh = ss.insertSheet(SHEET_EMAIL_LOG);
  const header = ['Timestamp', 'Advisor', 'Week', 'MoM Doc Link', 'CC Recipients', 'Sent By'];
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT);
  sh.getDataRange().setWrap(true);
  sh.setFrozenRows(1);
}

function logEmailSent(advisor, week, docUrl, ccList) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_EMAIL_LOG);
  sh.appendRow([new Date(), advisor, week, docUrl, ccList.join(', '), Session.getActiveUser().getEmail()]);
}