/**
 * ADVISOR ROSTER + STUDENT ROSTER
 * These two tabs are YOUR data entry - nothing here is pre-filled by
 * the script. Type advisor names/emails/divisions into Advisor Roster,
 * and students (with their advisor's name) into Student Roster.
 * Every other part of the system (Attendance, MoM Log, Profiles,
 * Reports) reads live from these two tabs - there is no hardcoded
 * list anywhere else to keep in sync.
 *
 * Layout (same on both tabs): row 1 = title band, row 2 = column
 * headers, row 3+ = your data.
 */
const ROSTER_HEADER_ROW = 2;
const ROSTER_DATA_START_ROW = 3;

/** Reads the Advisor Roster tab into [{name, email, division}, ...],
 *  skipping blank rows. This is what every other function calls
 *  instead of a hardcoded list. */
function getAdvisorList() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  if (!sh || sh.getLastRow() < ROSTER_DATA_START_ROW) return [];
  const data = sh.getRange(ROSTER_DATA_START_ROW, 1, sh.getLastRow() - ROSTER_DATA_START_ROW + 1, 3).getValues();
  return data.filter(r => r[0]).map(r => ({ name: r[0], email: r[1] || '', division: r[2] || '' }));
}

/** Reads the Student Roster tab into
 *  [[id, name, class, advisorName, subjectsOnProbation], ...],
 *  skipping blank rows. */
function getStudentList() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_STUDENT_ROSTER);
  if (!sh || sh.getLastRow() < ROSTER_DATA_START_ROW) return [];
  const data = sh.getRange(ROSTER_DATA_START_ROW, 1, sh.getLastRow() - ROSTER_DATA_START_ROW + 1, 5).getValues();
  return data.filter(r => r[1]).map(r => [r[0] || '', r[1], r[2] || '', r[3] || '', r[4] || '']);
}

/** Menu item #2: makes sure Advisor Roster's computed column (# Students
 *  Assigned) and the Attendance / MoM Log tabs match whoever is
 *  currently typed into Advisor Roster. Never touches names, emails,
 *  or anything else you've entered - purely a "keep everything else in
 *  sync with the roster" step. Run it any time after adding/removing
 *  an advisor or a batch of students. */
function generateAdvisorRoster() {
  const ss = SpreadsheetApp.getActive();
  setupAdvisorRosterSheet(ss);   // creates it blank if it doesn't exist yet - never touches it if it does
  setupStudentRosterSheet(ss);

  refreshStudentCounts(ss);

  const advisors = getAdvisorList();
  if (!advisors.length) {
    SpreadsheetApp.getUi().alert(
      'Advisor Roster is still empty. Add your advisors\' Name / Email / Division there first, ' +
      'then run this again to build Attendance and MoM Log from them.'
    );
    return;
  }

  rebuildAttendanceSheet(ss, advisors);
  rebuildMoMLogSheet(ss, advisors);

  SpreadsheetApp.getUi().alert(
    'Synced: Attendance and MoM Log now match your ' + advisors.length + ' advisor(s). ' +
    'Student counts refreshed from Student Roster (' + getStudentList().length + ' students).'
  );
}

/** Builds the blank Advisor Roster shell (title + headers only) the
 *  first time - never rebuilds it if it already exists, so your data
 *  is always safe. */
function setupAdvisorRosterSheet(ss) {
  let sh = ss.getSheetByName(SHEET_ROSTER);
  if (sh) return;
  sh = ss.insertSheet(SHEET_ROSTER);

  safeMerge_(sh, 'A1:J1').setValue('Advisor Roster')
    .setFontSize(15).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  sh.setRowHeight(1, 32);

  const header = [
    'Advisor Name', 'Email', 'Division', '# Students Assigned',
    'Weekly Follow-Up Day', 'Weekly Follow-Up Deadline',
    'Reports Required / Trimester', 'Week 3 Report Link',
    'Week 6 Report Link', 'Week 9 Report Link'
  ];
  sh.getRange(ROSTER_HEADER_ROW, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_TABLE_HEAD);

  // A handful of blank, ready-to-fill rows with the student-count formula pre-wired
  const blankRows = 15;
  for (let i = 0; i < blankRows; i++) {
    const r = ROSTER_DATA_START_ROW + i;
    sh.getRange(r, 4).setFormula("=COUNTIF('" + SHEET_STUDENT_ROSTER + "'!D:D, A" + r + ")");
    sh.getRange(r, 7).setValue(3);
  }

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(ROSTER_HEADER_ROW);
}

/** Builds the blank Student Roster shell the first time. Advisor Name
 *  is a dropdown that reads live from Advisor Roster's Name column, so
 *  it always matches whoever's actually on the roster - no separate
 *  list to keep in sync. */
function setupStudentRosterSheet(ss) {
  let sh = ss.getSheetByName(SHEET_STUDENT_ROSTER);
  if (sh) return;
  sh = ss.insertSheet(SHEET_STUDENT_ROSTER);

  safeMerge_(sh, 'A1:E1').setValue('Student Roster')
    .setFontSize(15).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  sh.setRowHeight(1, 32);

  const header = ['Student ID', 'Student Name', 'Class', 'Advisor Name', 'Subjects on Probation'];
  sh.getRange(ROSTER_HEADER_ROW, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_TABLE_HEAD);

  const rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  const blankRows = 60;
  if (rosterSheet) {
    const advisorNameRange = rosterSheet.getRange(ROSTER_DATA_START_ROW, 1, 200, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(advisorNameRange, true).setAllowInvalid(true).build();
    sh.getRange(ROSTER_DATA_START_ROW, 4, blankRows, 1).setDataValidation(rule);
  }

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(ROSTER_HEADER_ROW);
}

/** Rewrites the "# Students Assigned" formula down to match however
 *  many advisor rows currently exist (harmless to re-run any time). */
function refreshStudentCounts(ss) {
  const sh = ss.getSheetByName(SHEET_ROSTER);
  if (!sh) return;
  const lastRow = Math.max(sh.getLastRow(), ROSTER_DATA_START_ROW + 14);
  for (let r = ROSTER_DATA_START_ROW; r <= lastRow; r++) {
    const cell = sh.getRange(r, 4);
    if (!cell.getFormula()) cell.setFormula("=COUNTIF('" + SHEET_STUDENT_ROSTER + "'!D:D, A" + r + ")");
  }
}

/**
 * Menu item: clears every advisor row from Advisor Roster (and every
 * student row from Student Roster), keeping the title/header and
 * formatting intact. Use this once at the start of a new academic year
 * when the advisor list is starting fresh. Does NOT touch Settings,
 * Attendance, MoM Log, or anything else - only the two roster tabs.
 */
function resetAdvisorRoster() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Reset Advisor Roster',
    'This clears every advisor row from Advisor Roster and every student row from Student Roster ' +
    '(titles/headers stay). Use this at the start of a new year. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const roster = ss.getSheetByName(SHEET_ROSTER);
  if (roster && roster.getLastRow() >= ROSTER_DATA_START_ROW) {
    roster.getRange(ROSTER_DATA_START_ROW, 1, roster.getLastRow() - ROSTER_DATA_START_ROW + 1, roster.getLastColumn())
      .clearContent();
    refreshStudentCounts(ss);
  }
  const students = ss.getSheetByName(SHEET_STUDENT_ROSTER);
  if (students && students.getLastRow() >= ROSTER_DATA_START_ROW) {
    students.getRange(ROSTER_DATA_START_ROW, 1, students.getLastRow() - ROSTER_DATA_START_ROW + 1, students.getLastColumn())
      .clearContent();
  }
  ui.alert('Cleared. Type in this year\'s advisors and students, then run "Generate Advisor Roster."');
}

/** Helper: returns the row array [name,email,division,...] for an
 *  advisor from the Roster tab. */
function getRosterRow(advisorName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  for (let i = ROSTER_HEADER_ROW; i < data.length; i++) {
    if (data[i][0] === advisorName) return { rowNum: i + 1, values: data[i] };
  }
  return null;
}

/** Captures whatever you've already filled into the Roster tab
 *  (email, follow-up day/deadline, report links), keyed by advisor
 *  name, before it gets rebuilt. Currently unused since Roster is
 *  never auto-rebuilt once it exists, but kept for safety if you
 *  ever choose to reset it. */
function captureRosterState_(sheet) {
  const map = {};
  const data = sheet.getDataRange().getValues();
  const header = data[ROSTER_HEADER_ROW - 1];
  const idx = {
    email: header.indexOf('Email'),
    followDay: header.indexOf('Weekly Follow-Up Day'),
    followDeadline: header.indexOf('Weekly Follow-Up Deadline'),
    week3Link: header.indexOf('Week 3 Report Link'),
    week6Link: header.indexOf('Week 6 Report Link'),
    week9Link: header.indexOf('Week 9 Report Link'),
  };
  for (let r = ROSTER_HEADER_ROW; r < data.length; r++) {
    const name = data[r][0];
    if (!name) continue;
    map[name] = {
      email: idx.email !== -1 ? data[r][idx.email] : '',
      followDay: idx.followDay !== -1 ? data[r][idx.followDay] : '',
      followDeadline: idx.followDeadline !== -1 ? data[r][idx.followDeadline] : '',
      week3Link: idx.week3Link !== -1 ? data[r][idx.week3Link] : '',
      week6Link: idx.week6Link !== -1 ? data[r][idx.week6Link] : '',
      week9Link: idx.week9Link !== -1 ? data[r][idx.week9Link] : '',
    };
  }
  return map;
}