/**
 * TRIMESTER / WEEK REFERENCE
 * SAMPLE calendar — replace these dates with your own school calendar.
 * Weeks 1-10 are advisory meeting weeks each term.
 * Weeks 11/12 are Wrap Up & Exams — no advisory meetings; shown for context.
 * Use the note column for holidays or events that affect a week.
 */
function getTrimesterReferenceData() {
  return [
    // term, week, startDate, endDate, note
    [1, 1,  '2026-09-13', '2026-09-19', ''],
    [1, 2,  '2026-09-20', '2026-09-26', ''],
    [1, 3,  '2026-09-27', '2026-10-03', ''],
    [1, 4,  '2026-10-04', '2026-10-10', ''],
    [1, 5,  '2026-10-11', '2026-10-17', ''],
    [1, 6,  '2026-10-18', '2026-10-24', ''],
    [1, 7,  '2026-10-25', '2026-10-31', ''],
    [1, 8,  '2026-11-01', '2026-11-07', ''],
    [1, 9,  '2026-11-08', '2026-11-14', ''],
    [1, 10, '2026-11-15', '2026-11-21', ''],
    [1, 11, '2026-11-22', '2026-11-28', 'WRAP UP & EXAMS - no advisory meeting'],
    [1, 12, '2026-11-29', '2026-12-05', 'WRAP UP & EXAMS - no advisory meeting'],
    [2, 1,  '2026-12-07', '2026-12-12', ''],
    [2, 2,  '2026-12-13', '2026-12-19', ''],
    [2, 3,  '2026-12-20', '2026-12-26', ''],
    [2, 4,  '2027-01-10', '2027-01-16', ''],
    [2, 5,  '2027-01-17', '2027-01-23', ''],
    [2, 6,  '2027-01-24', '2027-01-30', ''],
    [2, 7,  '2027-01-31', '2027-02-06', ''],
    [2, 8,  '2027-02-07', '2027-02-13', ''],
    [2, 9,  '2027-02-14', '2027-02-20', ''],
    [2, 10, '2027-02-21', '2027-02-27', ''],
    [2, 11, '2027-03-14', '2027-03-20', 'WRAP UP & EXAMS - no advisory meeting'],
    [2, 12, '2027-03-21', '2027-03-27', 'WRAP UP & EXAMS - no advisory meeting'],
    [3, 1,  '2027-03-29', '2027-04-03', ''],
    [3, 2,  '2027-04-04', '2027-04-10', ''],
    [3, 3,  '2027-04-11', '2027-04-17', ''],
    [3, 4,  '2027-04-18', '2027-04-24', ''],
    [3, 5,  '2027-04-25', '2027-05-01', ''],
    [3, 6,  '2027-05-02', '2027-05-08', ''],
    [3, 7,  '2027-05-09', '2027-05-15', ''],
    [3, 8,  '2027-05-23', '2027-05-29', ''],
    [3, 9,  '2027-05-30', '2027-06-05', ''],
    [3, 10, '2027-06-07', '2027-06-12', ''],
    [3, 11, '2027-06-13', '2027-06-19', 'WRAP UP & EXAMS - no advisory meeting'],
    [3, 12, '2027-06-20', '2027-06-26', 'WRAP UP & EXAMS - no advisory meeting'],
  ];
}

function setupTrimesterReferenceSheet(ss) {
  let sh = ss.getSheetByName(SHEET_TRIMESTER);
  if (sh) ss.deleteSheet(sh); // always rebuild from source data - this tab is read-only reference
  sh = ss.insertSheet(SHEET_TRIMESTER);

  const header = ['Term', 'Week', 'Start Date', 'End Date', 'Notes / Day Off / Event'];
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT);

  const data = getTrimesterReferenceData().map(r => [
    r[0], r[1], new Date(r[2]), new Date(r[3]), r[4]
  ]);
  sh.getRange(2, 1, data.length, 5).setValues(data);
  sh.getRange(2, 3, data.length, 2).setNumberFormat('ddd, d mmm yyyy');

  // Shade Wrap Up & Exams weeks (11/12) and highlight rows with a note
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    if (data[i][4].toString().indexOf('WRAP UP') === 0) {
      sh.getRange(rowNum, 1, 1, 5).setBackground('#f4cccc');
    } else if (data[i][4]) {
      sh.getRange(rowNum, 1, 1, 5).setBackground('#fff2cc');
    }
  }

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(1);
  protectReadOnly(sh, 'Reference only - update via source calendar, not by hand.');
}

function protectReadOnly(sheet, note) {
  const protection = sheet.protect().setDescription(note || 'Read-only');
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

/** Given today's date, returns {term, week} or null if in a break/exam gap. */
function getCurrentTermWeek() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const rows = getTrimesterReferenceData();
  for (const r of rows) {
    const start = new Date(r[2]);
    const end = new Date(r[3]);
    end.setHours(23, 59, 59, 999);
    if (today >= start && today <= end) {
      return { term: r[0], week: r[1], note: r[4] };
    }
  }
  return null;
}

/** Returns the Start Date (as a Date, or '' if not found) for a
 *  given term/week - used to show "This Week's Meeting Date" on
 *  profile/week tabs without the advisor having to set anything. */
function getDateForTermWeek(term, week) {
  const rows = getTrimesterReferenceData();
  const match = rows.find(r => r[0] === term && r[1] === week);
  return match ? new Date(match[2]) : '';
}

/** Same as getDateForTermWeek, but shifted to whichever weekday
 *  matches the advisor's Weekly Follow-Up Day (e.g. Sunday,
 *  Thursday) instead of always the week's generic Sunday start -
 *  used for "This Week's Meeting Date" so it shows the day they
 *  actually meet, not just the first day of the school week. */
function getMeetingDateForWeek(term, week, followUpDayName) {
  const weekStart = getDateForTermWeek(term, week);
  if (!weekStart) return null;
  const dayOffsets = {
    sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6,
  };
  const offset = dayOffsets[String(followUpDayName || '').trim().toLowerCase()];
  const d = new Date(weekStart);
  if (offset !== undefined) d.setDate(d.getDate() + offset);
  return d;
}