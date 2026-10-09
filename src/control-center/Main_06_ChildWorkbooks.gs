/**
 * Creates (or refreshes) one advisor's own Google Sheet workbook:
 * a Dashboard, a hidden Config tab (used by that workbook's own
 * script to reach back into this main workbook - see Child_Code.gs),
 * and one tab per advisory week (1-10). Weeks 3, 6 and 9 also get a
 * "Report - Week N" tab where the advisor fills a per-student table;
 * once complete, their own workbook auto-generates one report PDF
 * per student and pushes the folder link back here - see
 * Child_Code.gs's finalizeReport().
 *
 * Requires "Advisor Workbook Template Spreadsheet ID" in Settings:
 * a blank Google Sheet that has Child_Code.gs pasted into its own
 * Extensions > Apps Script. See README "One-time template setup".
 */
function createOrUpdateAdvisorWorkbook(advisorName, folder, mainId) {
  const templateId = getSetting('Advisor Workbook Template Spreadsheet ID');
  if (!templateId) {
    throw new Error('Set "Advisor Workbook Template Spreadsheet ID" in Settings first (see README).');
  }

  const fileName = 'Advisor Workbook - ' + advisorName;
  const existing = folder.getFilesByName(fileName);
  let file;
  if (existing.hasNext()) {
    file = existing.next(); // refresh in place rather than duplicating
  } else {
    file = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
  }

  const childSs = SpreadsheetApp.openById(file.getId());

  // Config tab (hidden) - tells the child's own script where "home" is
  let cfg = childSs.getSheetByName('Config');
  if (!cfg) cfg = childSs.insertSheet('Config');
  cfg.clear();
  cfg.getRange('A1:B3').setValues([
    ['MAIN_WORKBOOK_ID', mainId],
    ['ADVISOR_NAME', advisorName],
    ['ROSTER_SHEET', SHEET_ROSTER],
  ]);
  cfg.hideSheet();

  // Dashboard tab
  let dash = childSs.getSheetByName('Dashboard');
  if (!dash) dash = childSs.insertSheet('Dashboard', 0);
  dash.clear();
  safeMerge_(dash, 'A1:D1').setValue('Welcome, ' + advisorName)
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  dash.setRowHeight(1, 32);
  dash.getRange('A3').setValue('Today:').setFontWeight('bold');
  dash.getRange('B3').setValue(new Date()).setNumberFormat('dddd, d mmmm yyyy');
  const tw = getCurrentTermWeek();
  dash.getRange('A4').setValue('Current Week:').setFontWeight('bold');
  dash.getRange('B4').setValue(tw ? 'Term ' + tw.term + ' - Week ' + tw.week : 'Break / Exam period');
  dash.getRange('A3:B4').setBackground(COLOR_READONLY_BG);
  dash.getRange('A6').setValue(
    'Fill in your weekly tab (left) before your meeting with the Academic Dean. ' +
    'On Week 3, 6 and 9, fill in the "Report - Week N" tab for every student - once ' +
    'everything is filled in, you\'ll be asked to finalize and it handles the rest.'
  ).setWrap(true).setBackground(COLOR_SECTION_BG).setFontColor(COLOR_SECTION_TXT);
  safeMerge_(dash, 'A6:D6');
  dash.getRange('A8').setValue('First time opening this workbook? Run:');
  dash.getRange('A9').setValue('Extensions > Apps Script > run "enableAutoSync" once, or use the "Advisor" menu above.');
  dash.getDataRange().setWrap(true);

  // Week tabs
  const weekCount = getWeekCount();
  for (let w = 1; w <= weekCount; w++) {
    const tabName = 'Week ' + w;
    let sh = childSs.getSheetByName(tabName);
    const preserved = sh ? captureEditableState(sh) : null;
    if (!sh) sh = childSs.insertSheet(tabName);
    writeProfileLayout(sh, advisorName, w, /*isMain=*/false);
    if (preserved) restoreEditableState(sh, preserved);

    if (w === 3 || w === 6 || w === 9) {
      const reportLabel = 'Week ' + w + ' Reports Folder Link';
      const existingFolderLink = sh ? readFolderLinkLabel_(sh, reportLabel) : '';
      buildReportTemplateTab(childSs, advisorName, w);
      const lastRow = sh.getLastRow() + 2;
      sh.getRange(lastRow, 1).setValue(reportLabel + ' (fills automatically once finalized):').setFontWeight('bold').setWrap(true);
      sh.getRange(lastRow, 2).setValue(existingFolderLink || '');
    }
  }

  // remove default "Sheet1" if the template left one behind
  const blank = childSs.getSheetByName('Sheet1');
  if (blank && childSs.getSheets().length > 1) childSs.deleteSheet(blank);
}

function readFolderLinkLabel_(sheet, label) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return '';
  const data = sheet.getRange(1, 1, lastRow, 2).getValues();
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).indexOf(label) === 0) return data[r][1];
  }
  return '';
}

/**
 * Builds the simple fill-in report table advisors complete every
 * 3 weeks: one row per student, a Better/Same/Worse trend dropdown,
 * two short text fields, and a Report Link column that fills itself
 * once the report is finalized. Regenerating never wipes what's
 * already been filled in - captured before rebuild, restored after.
 */
function buildReportTemplateTab(childSs, advisorName, week) {
  const tabName = 'Report - Week ' + week;
  let sh = childSs.getSheetByName(tabName);
  const preserved = sh ? captureReportState_(sh) : null;
  if (!sh) sh = childSs.insertSheet(tabName);
  sh.clear();
  sh.clearConditionalFormatRules();
  clearAllValidations_(sh);

  safeMerge_(sh, 'A1:E1').setValue('Advisor 3-Week Report')
    .setFontSize(14).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_HEADER_BG);
  sh.setRowHeight(1, 30);
  sh.getRange('A2').setValue('Advisor: ' + advisorName + '   |   Week ' + week).setFontWeight('bold');
  sh.getRange('A3').setValue(
    'Fill in every student\'s Trend, Key Concern and Recommendation. Once all rows are complete, ' +
    'a prompt appears asking you to finalize - click Yes and the reports generate automatically.'
  ).setFontStyle('italic').setWrap(true);
  safeMerge_(sh, 'A3:E3');
  sh.getRange('A4').setValue('Status:').setFontWeight('bold');
  sh.getRange('B4').setValue(preserved ? preserved.status || 'Not Finalized' : 'Not Finalized');
  sh.getRange('A4:B4').setBackground(COLOR_READONLY_BG).setFontColor(COLOR_READONLY_TXT);
  addStatusColorRule(sh, sh.getRange('B4'), { 'Finalized': STATUS_COMPLETED, 'Not Finalized': STATUS_NOT_STARTED });

  const header = ['Student Name', 'Trend', 'Key Concern', 'Recommendation', 'Report Link'];
  sh.getRange(5, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT).setBackground(COLOR_TABLE_HEAD);

  const students = getStudentList().filter(s => s[3] === advisorName);
  if (students.length) {
    const rows = students.map(s => {
      const prev = preserved && preserved.students[s[1]];
      return [s[1], prev ? prev.trend : '', prev ? prev.concern : '', prev ? prev.recommendation : '', prev ? prev.link : ''];
    });
    sh.getRange(6, 1, rows.length, header.length).setValues(rows);
    const trendRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Better', 'Same', 'Worse'], true).setAllowInvalid(false).build();
    sh.getRange(6, 2, rows.length, 1).setDataValidation(trendRule);
    addStatusColorRule(sh, sh.getRange(6, 2, rows.length, 1), {
      'Better': STATUS_COMPLETED, 'Same': STATUS_IN_PROGRESS, 'Worse': STATUS_OVERDUE,
    });
    for (let i = 0; i < rows.length; i++) {
      if (i % 2 === 1) sh.getRange(6 + i, 1, 1, header.length).setBackground(COLOR_ALT_ROW);
    }
    sh.getRange(6, 5, rows.length, 1).setFontColor(COLOR_READONLY_TXT); // Report Link is auto-filled, not typed
  }

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(5);
}

/** Captures whatever an advisor already filled in on a Report tab
 *  (status, trend/concern/recommendation/link per student), keyed
 *  by student name, before the tab gets rebuilt. */
function captureReportState_(sheet) {
  const state = { status: '', students: {} };
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return state;
  state.status = sheet.getRange('B4').getValue();
  if (lastRow < 6) return state;
  const data = sheet.getRange(6, 1, lastRow - 5, 5).getValues();
  data.forEach(r => {
    if (!r[0]) return;
    state.students[r[0]] = { trend: r[1], concern: r[2], recommendation: r[3], link: r[4] };
  });
  return state;
}