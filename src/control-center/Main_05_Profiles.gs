/**
 * VISUAL THEME - mapped to your Google Sheets theme spec.
 * Swap these hex codes any time - every tab pulls from here.
 */
// PRIMARY
const COLOR_HEADER_BG   = '#1F3864'; // Dark Navy - main titles / main headers
const COLOR_HEADER_TEXT = '#FFFFFF';
const COLOR_TABLE_HEAD  = '#2E5C8A'; // Dark Blue - secondary headers (student/report table headers)
const COLOR_SECTION_BG  = '#CFE2F3'; // Light Blue - section headers
const COLOR_SECTION_TXT = '#1F3864';
const COLOR_KPI_TXT     = '#0B5E5A'; // Dark Teal - KPIs / key highlights
const COLOR_KPI_BG      = '#D0F0EE'; // Light Teal - KPI backgrounds / highlights
const COLOR_WHITE       = '#FFFFFF'; // main working area
const COLOR_READONLY_BG = '#F1F3F4'; // Light Gray - separators / inactive (Dean-only fields)
const COLOR_READONLY_TXT= '#3C3C3C'; // Dark Gray - main text
const COLOR_ALT_ROW     = '#F1F3F4'; // Light Gray - zebra banding

// STATUS
const STATUS_COMPLETED   = '#D9EAD3'; // Light Green
const STATUS_IN_PROGRESS = '#FFF2CC'; // Light Yellow
const STATUS_AT_RISK     = '#FCE5CD'; // Light Orange
const STATUS_OVERDUE     = '#F4CCCC'; // Light Red
const STATUS_PENDING     = '#E6D6F5'; // Light Purple
const STATUS_NOT_STARTED = '#F1F3F4'; // Light Gray

// FUNCTIONAL
const FUNC_INFO     = '#CFE2F3'; // Light Blue - information
const FUNC_NOTES    = '#FFF2CC'; // Light Yellow - notes / comments
const FUNC_PRIORITY = '#0B5E5A'; // Dark Teal - important / priority
const FUNC_STANDARD = '#FFFFFF'; // standard cells

/**
 * Checklist items pulled from Advisor_Checklist.docx.
 * The FC/ET/Quiz/HW/VP/CW/AR/Attendance/Counsellor items are now
 * individual columns in the per-student table (see
 * writeStudentChecklistTable) rather than a flat list, since they
 * are checked per student, not once for the week.
 */
function getDailyChecklistItems() {
  return [
    'Check in briefly with each of your assigned students, even a quick hello.',
    'Notice and note anything unusual - mood, engagement, appearance, friendships.',
    "Confirm they know what's due today in the subject(s) they're on probation in.",
  ];
}
function getBringToMeetingItems() {
  return [
    'Updated FC / ET / Quizzes / HW / VP / CW / AR status for every assigned student.',
    "Attendance concerns - who's missing sessions, and how often.",
    'Any new academic, behavioral, or wellness concerns.',
    'Any improvements or wins worth recording.',
  ];
}
function getReportImmediatelyItems() {
  return [
    'Signs of serious emotional distress.',
    'A sudden drop in engagement, or repeated unexplained absences.',
    'Any disclosure that raises a safety concern.',
    'Repeated missed assignments despite your check-ins.',
  ];
}

/** Menu item #3: builds one Profile tab per advisor in the MAIN workbook,
 *  then creates/updates each advisor's own mirrored workbook.
 *  Safe to re-run any time - advisor-entered data (checkboxes, statuses,
 *  notes) is captured before rebuild and restored after, so nothing
 *  they've already filled in is ever lost. */
function generateAdvisorProfiles() {
  const ss = SpreadsheetApp.getActive();
  const advisors = getAdvisorList();
  const folderId = getSetting('Folder Term ID (advisor workbooks live here)');
  const folder = folderId ? DriveApp.getFolderById(folderId) : null;

  advisors.forEach(a => {
    // 1. Profile tab inside the main workbook
    const tabName = PROFILE_PREFIX + a.name;
    const existingSheet = ss.getSheetByName(tabName);
    const preserved = existingSheet ? captureEditableState(existingSheet) : null;
    let sh = existingSheet;
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet(tabName);
    writeProfileLayout(sh, a.name, null, /*isMain=*/true);
    if (preserved) restoreEditableState(sh, preserved);

    // 2. Advisor's own workbook (separate Sheet file)
    if (folder) {
      createOrUpdateAdvisorWorkbook(a.name, folder, ss.getId());
    }
  });

  SpreadsheetApp.getUi().alert(
    'Generated ' + advisors.length + ' advisor profile tabs' +
    (folder ? ' and their individual workbooks in the Folder Term ID.' :
      '. Set "Folder Term ID" in Settings and re-run to also create individual workbooks.') +
    ' Anything advisors had already filled in was preserved.'
  );
}

/**
 * Captures everything an ADVISOR (not the Dean) would have filled in
 * on a profile/week tab: the Confidentiality checkbox, Daily/Bring-to-
 * meeting checklist statuses, and every per-student editable field.
 * Called right before a tab gets rebuilt so re-running Generate is safe.
 */
function captureEditableState(sheet) {
  const state = { checklist: {}, confidentiality: null, students: {} };
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 17);
  if (lastRow < 1) return state;
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const trackedLabels = ['Confidentiality acknowledged for this week']
    .concat(getDailyChecklistItems()).concat(getBringToMeetingItems());

  for (let r = 0; r < data.length; r++) {
    const label = String(data[r][0]).trim();
    if (label === 'Confidentiality acknowledged for this week') state.confidentiality = data[r][1];
    else if (trackedLabels.indexOf(label) !== -1) state.checklist[label] = data[r][1];
  }

  const headerRowIdx = data.findIndex(r => r[0] === 'Student ID');
  if (headerRowIdx !== -1) {
    for (let r = headerRowIdx + 1; r < data.length; r++) {
      const name = data[r][1];
      if (!name) break;
      state.students[name] = {
        fc: data[r][4], et: data[r][5], quiz: data[r][6], hw: data[r][7],
        vp: data[r][8], cw: data[r][9], ar: data[r][10], attendance: data[r][11],
        weeklyStatus: data[r][12], why: data[r][13], flagCounsellor: data[r][14],
        onTrack: data[r][15], notes: data[r][16],
      };
    }
  }
  return state;
}

/** Writes a captured state (see captureEditableState) back onto a
 *  freshly rebuilt sheet, matching students by name. */
function restoreEditableState(sheet, state) {
  if (!state) return;
  if (state.confidentiality !== null) {
    setLabeledCell(sheet, 'Confidentiality acknowledged for this week', state.confidentiality);
  }
  Object.keys(state.checklist).forEach(label => setLabeledCell(sheet, label, state.checklist[label]));

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 17);
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRowIdx = data.findIndex(r => r[0] === 'Student ID');
  if (headerRowIdx === -1) return;

  for (let r = headerRowIdx + 1; r < data.length; r++) {
    const name = data[r][1];
    if (!name) break;
    const s = state.students[name];
    if (!s) continue;
    const rowNum = r + 1;
    sheet.getRange(rowNum, 5, 1, 8).setValues([[s.fc, s.et, s.quiz, s.hw, s.vp, s.cw, s.ar, s.attendance]]);
    sheet.getRange(rowNum, 13).setValue(s.weeklyStatus);
    sheet.getRange(rowNum, 14).setValue(s.why);
    sheet.getRange(rowNum, 15).setValue(s.flagCounsellor);
    sheet.getRange(rowNum, 16).setValue(s.onTrack);
    sheet.getRange(rowNum, 17).setValue(s.notes);
  }
}

/**
 * Shared layout writer - used for both the main-workbook profile tab
 * and every "Week N" tab inside each advisor's own workbook.
 * @param week   null for the main tab (uses current term/week);
 *               a number 1-10 for a specific child "Week N" tab.
 * @param isMain true for the main workbook (fields stay editable there
 *               since you're the one who sets them); false for a
 *               child workbook (Dean-controlled fields get locked).
 */
function writeProfileLayout(sh, advisorName, week, isMain) {
  sh.clear();
  sh.clearFormats();
  sh.clearConditionalFormatRules();
  clearAllValidations_(sh);
  getProtections_(sh).forEach(p => { try { p.remove(); } catch (e) {} });

  const advisorInfo = getAdvisorList().find(a => a.name === advisorName) || {};
  const roster = getRosterRow(advisorName);
  const followDay = roster ? roster.values[4] : '';
  const followDeadline = roster ? roster.values[5] : '';

  const tw = week ? { term: Number(getSetting('Term # (1, 2, or 3)')) || 1, week: week } : getCurrentTermWeek();
  const meetingDate = tw ? getMeetingDateForWeek(tw.term, tw.week, followDay) : null;

  // ---- Title band ----
  safeMerge_(sh, 'A1:Q1').setValue('Advisor Profile - ' + advisorName)
    .setFontSize(15).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT)
    .setBackground(COLOR_HEADER_BG).setHorizontalAlignment('left');
  sh.setRowHeight(1, 32);

  // ---- Dean-mirrored info block (rows 2-7) - read only in child workbooks ----
  const infoRows = [
    ['Division', advisorInfo.division || ''],
    ['Weekly Follow-Up Day', followDay || '(set in Roster tab)'],
    ['Weekly Follow-Up Deadline', followDeadline || '(set in Roster tab)'],
    ['This Week\u2019s Meeting Date', meetingDate ? meetingDate : '(no active advisory week)'],
    ['Attend Weekly Meeting', 'Not Done'],
    ['MoM Received', 'Not Done'],
  ];
  sh.getRange(2, 1, infoRows.length, 2).setValues(infoRows);
  sh.getRange(2, 1, infoRows.length, 1).setFontWeight('bold');
  const infoBlockRange = sh.getRange(2, 1, infoRows.length, 2);
  infoBlockRange.setBackground(COLOR_READONLY_BG).setFontColor(COLOR_READONLY_TXT);
  if (!isMain) protectRangeOwnerOnly_(infoBlockRange, 'Set by the Academic Dean - read only');

  let row = infoRows.length + 3;

  // ---- Confidentiality (advisor sets this weekly) ----
  sh.getRange(row, 1).setValue('Confidentiality acknowledged for this week').setFontWeight('bold');
  sh.getRange(row, 2).insertCheckboxes();
  row += 2;

  // ---- Daily checklist ----
  row = writeChecklistBlock(sh, row, 'Daily Checklist', getDailyChecklistItems());
  row += 1;

  // ---- What to bring / report immediately (reference, amber banner) ----
  row = writeDisclaimerBlock(sh, row, 'What to Bring to Your Weekly Meeting with the Dean', getBringToMeetingItems(), true);
  row += 1;
  row = writeDisclaimerBlock(sh, row, "Report Immediately - Don't Wait for the Weekly Meeting", getReportImmediatelyItems(), false);
  row += 2;

  // ---- Horizontal per-student checklist ----
  const students = getStudentList().filter(s => s[3] === advisorName);
  writeStudentChecklistTable(sh, row, students, isMain);

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(1);
}

function writeChecklistBlock(sh, startRow, title, items) {
  let row = startRow;
  sh.getRange(row, 1).setValue(title).setFontWeight('bold')
    .setBackground(COLOR_SECTION_BG).setFontColor(COLOR_SECTION_TXT);
  sh.getRange(row, 1, 1, 2).setBackground(COLOR_SECTION_BG);
  row += 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Done', 'Not Done'], true).setAllowInvalid(false).build();
  const firstDataRow = row;
  items.forEach(item => {
    sh.getRange(row, 1).setValue(item);
    sh.getRange(row, 2).setValue('Not Done').setDataValidation(rule);
    row += 1;
  });
  if (items.length) {
    addStatusColorRule(sh, sh.getRange(firstDataRow, 2, items.length, 1), {
      'Done': STATUS_COMPLETED, 'Not Done': STATUS_NOT_STARTED,
    });
  }
  return row;
}

function writeDisclaimerBlock(sh, startRow, title, items, withCheckboxes) {
  let row = startRow;
  sh.getRange(row, 1).setValue(title).setFontWeight('bold').setFontStyle('italic')
    .setBackground(COLOR_SECTION_BG).setFontColor(COLOR_SECTION_TXT);
  sh.getRange(row, 1, 1, 2).setBackground(COLOR_SECTION_BG);
  row += 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Done', 'Not Done'], true).setAllowInvalid(false).build();
  const firstDataRow = row;
  items.forEach(item => {
    sh.getRange(row, 1).setValue('\u2022 ' + item);
    if (withCheckboxes) sh.getRange(row, 2).setValue('Not Done').setDataValidation(rule);
    row += 1;
  });
  if (withCheckboxes && items.length) {
    addStatusColorRule(sh, sh.getRange(firstDataRow, 2, items.length, 1), {
      'Done': STATUS_COMPLETED, 'Not Done': STATUS_NOT_STARTED,
    });
  }
  return row;
}

/**
 * The horizontal per-student checklist table:
 * ID | Name | Class | Subjects | FC | ET | Quiz | HW | VP | CW | AR |
 * Attendance | Weekly Status | Why | Flag Counsellor? | On Track? | Notes
 */
function writeStudentChecklistTable(sh, startRow, students, isMain) {
  sh.getRange(startRow, 1).setValue('Students - Weekly Check').setFontWeight('bold')
    .setFontSize(12);
  const headerRow = startRow + 1;
  const header = [
    'Student ID', 'Student Name', 'Class', 'Subjects on Probation',
    'FC', 'ET', 'Quiz', 'HW', 'VP', 'CW', 'AR', 'Attendance',
    'Weekly Status', 'Why', 'Flag Counsellor?', 'On Track?', 'Notes (mandatory)'
  ];
  sh.getRange(headerRow, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_TABLE_HEAD)
    .setWrap(true);

  if (!students.length) {
    sh.getDataRange().setWrap(true);
    sh.setFrozenRows(headerRow);
    return;
  }

  const checkRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Better', 'Same', 'Worse'], true).setAllowInvalid(false).build();
  const trackRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['On Track', 'Not on Track'], true).setAllowInvalid(false).build();

  const dataRow = headerRow + 1;
  const rows = students.map(s => [
    s[0], s[1], s[2], s[4] || '',
    false, false, false, false, false, false, false, false,
    '', '', false, '', ''
  ]);
  sh.getRange(dataRow, 1, rows.length, header.length).setValues(rows);

  sh.getRange(dataRow, 5, rows.length, 8).setDataValidation(checkRule);   // FC..Attendance
  sh.getRange(dataRow, 13, rows.length, 1).setDataValidation(statusRule); // Weekly Status
  sh.getRange(dataRow, 15, rows.length, 1).setDataValidation(checkRule);  // Flag Counsellor
  sh.getRange(dataRow, 16, rows.length, 1).setDataValidation(trackRule);  // On Track

  // Zebra striping for readability
  for (let i = 0; i < rows.length; i++) {
    if (i % 2 === 1) sh.getRange(dataRow + i, 1, 1, header.length).setBackground(COLOR_ALT_ROW);
  }

  // Status coloring: completed checks go green, flags go orange,
  // trend/track columns traffic-light so you can scan a whole roster fast
  addBooleanTrueColorRule(sh, sh.getRange(dataRow, 5, rows.length, 8), STATUS_COMPLETED);
  addBooleanTrueColorRule(sh, sh.getRange(dataRow, 15, rows.length, 1), STATUS_AT_RISK);
  addStatusColorRule(sh, sh.getRange(dataRow, 13, rows.length, 1), {
    'Better': STATUS_COMPLETED, 'Same': STATUS_IN_PROGRESS, 'Worse': STATUS_OVERDUE,
  });
  addStatusColorRule(sh, sh.getRange(dataRow, 16, rows.length, 1), {
    'On Track': STATUS_COMPLETED, 'Not on Track': STATUS_OVERDUE,
  });
  sh.getRange(dataRow, 17, rows.length, 1).setBackground(FUNC_NOTES); // Notes - functional "notes" color

  // ID, Name, Class, Subjects are Dean-controlled -> read-only for advisors
  const idNameClassSubjects = sh.getRange(dataRow, 1, rows.length, 4);
  idNameClassSubjects.setFontColor(COLOR_READONLY_TXT);
  if (!isMain) protectRangeOwnerOnly_(idNameClassSubjects, 'Set by the Academic Dean - read only');

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(headerRow);
}

/** Restricts a range to edits by the file's owner only (used for the
 *  fields you control - meeting date, attendance, subjects, etc). */
function protectRangeOwnerOnly_(range, description) {
  const protection = range.protect().setDescription(description || 'Read only');
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function getProtections_(sheet) {
  return sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
}

/** Looks up a label in column A (top info block or checklist rows)
 *  and sets the value in column B of that same row. */
function setLabeledCell(sheet, label, value) {
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === label) {
      sheet.getRange(r + 1, 2).setValue(value);
      return true;
    }
  }
  return false;
}

/** Colors a text/dropdown range by its exact value (Better->green,
 *  Worse->red, etc). valueColorMap = {textValue: hexColor}. */
function addStatusColorRule(sheet, range, valueColorMap) {
  const rules = sheet.getConditionalFormatRules();
  Object.keys(valueColorMap).forEach(val => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(val)
        .setBackground(valueColorMap[val])
        .setRanges([range])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
}

/** Colors a checkbox range green (or any color) only when TRUE -
 *  used so completed items visibly light up as advisors check them. */
function addBooleanTrueColorRule(sheet, range, color) {
  const firstCell = range.getCell(1, 1).getA1Notation();
  const rules = sheet.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + firstCell + '=TRUE')
      .setBackground(color)
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(rules);
}

/**
 * Menu item: builds a throwaway "TEST - Sandbox Advisor" profile tab
 * with 3 fake students, so you can click around, check boxes, and try
 * things without touching real advisor/student data. "Generate Advisor
 * Profiles" never looks at or deletes this tab since it isn't part of
 * your real advisor list. Re-run any time to reset it.
 */
function createTestAdvisorProfile() {
  const ss = SpreadsheetApp.getActive();
  const testName = 'TEST - Sandbox Advisor';
  const tabName = PROFILE_PREFIX + testName;
  let sh = ss.getSheetByName(tabName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(tabName);

  sh.getRange('A1:Q1').merge()
    .setValue('Advisor Profile - ' + testName + '  (safe to experiment - not real data)')
    .setFontSize(15).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  sh.setRowHeight(1, 32);

  const infoRows = [
    ['Division', 'TEST'],
    ['Weekly Follow-Up Day', 'N/A'],
    ['Weekly Follow-Up Deadline', 'N/A'],
    ['This Week\u2019s Meeting Date', new Date()],
    ['Attend Weekly Meeting', 'Not Done'],
    ['MoM Received', 'Not Done'],
  ];
  sh.getRange(2, 1, infoRows.length, 2).setValues(infoRows);
  sh.getRange(2, 1, infoRows.length, 1).setFontWeight('bold');
  sh.getRange(2, 1, infoRows.length, 2).setBackground(COLOR_READONLY_BG).setFontColor(COLOR_READONLY_TXT);

  let row = infoRows.length + 3;
  sh.getRange(row, 1).setValue('Confidentiality acknowledged for this week').setFontWeight('bold');
  sh.getRange(row, 2).insertCheckboxes();
  row += 2;
  row = writeChecklistBlock(sh, row, 'Daily Checklist', getDailyChecklistItems());
  row += 1;
  row = writeDisclaimerBlock(sh, row, 'What to Bring to Your Weekly Meeting with the Dean', getBringToMeetingItems(), true);
  row += 1;
  row = writeDisclaimerBlock(sh, row, "Report Immediately - Don't Wait for the Weekly Meeting", getReportImmediatelyItems(), false);
  row += 2;

  const fakeStudents = [
    ['T001', 'Test Student One', 'Grade 7'],
    ['T002', 'Test Student Two', 'Grade 8'],
    ['T003', 'Test Student Three', 'Grade 9'],
  ];
  writeStudentChecklistTable(sh, row, fakeStudents, /*isMain=*/true);
  sh.getDataRange().setWrap(true);
  sh.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'Test profile "' + tabName + '" created with 3 fake students. Click around freely - ' +
    '"Generate Advisor Profiles" will never touch this tab. Re-run this menu item any time to reset it.'
  );
}