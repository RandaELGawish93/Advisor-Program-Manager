/**
 * ============================================================
 * ADVISOR WORKBOOK CONTROL CENTER — MAIN WORKBOOK SCRIPT
 * ============================================================
 * This script is bound to your MAIN "Advisor Workbook
 * Control Center" Google Sheet. Paste this file (and the other
 * Main_*.gs files) into Extensions > Apps Script on that sheet.
 *
 * See README.md for the full setup sequence.
 */

// ---- SHEET NAME CONSTANTS ----
const SHEET_DASHBOARD       = 'Dashboard';
const SHEET_SETTINGS        = 'Settings';
const SHEET_TRIMESTER       = 'Trimester Reference';
const SHEET_ROSTER          = 'Advisor Roster';
const SHEET_STUDENT_ROSTER  = 'Student Roster';
const SHEET_ATTENDANCE      = 'Advisor Attendance';
const SHEET_MOM_LOG         = 'MoM Log';
const SHEET_EMAIL_LOG       = 'Email Log';
const PROFILE_PREFIX        = 'Profile - ';

// ---- MENU ----
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Advisor System')
    .addItem('1. Initialize Workbook (run once)', 'initializeWorkbook')
    .addSeparator()
    .addItem('2. Generate Advisor Roster', 'generateAdvisorRoster')
    .addItem('Reset Advisor Roster (new academic year)', 'resetAdvisorRoster')
    .addItem('3. Generate Advisor Profiles + Workbooks', 'generateAdvisorProfiles')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Enable Live Sync (Attendance → Profiles)', 'enableMainLiveSync')
    .addItem('Pull Live Status from Advisor Workbooks', 'pullLiveStatusFromAdvisors')
    .addItem('Send Weekly MoM for Selected Row', 'sendMoMForSelectedRow')
    .addSeparator()
    .addItem('Archive This Trimester', 'archiveTrimester')
    .addToUi();
}

/**
 * Run this ONCE per trimester setup (or once ever, then re-run
 * "Generate Advisor Roster/Profiles" each new trimester after
 * you update Settings and Trimester #).
 * Note: Advisor Roster and Student Roster are created here BLANK -
 * this is where you type in your advisors and students. Attendance
 * and MoM Log are NOT built here - they're built from "Generate
 * Advisor Roster" instead, reading whoever you've actually entered.
 */
function initializeWorkbook() {
  const ss = SpreadsheetApp.getActive();
  setupSettingsSheet(ss);
  setupTrimesterReferenceSheet(ss);
  setupDashboardSheet(ss);
  setupEmailLogSheet(ss);
  setupAdvisorRosterSheet(ss);
  setupStudentRosterSheet(ss);
  SpreadsheetApp.getUi().alert(
    'Workbook initialized. Now fill in the Settings tab, then type your advisors into ' +
    '"Advisor Roster" and your students into "Student Roster" (both blank, ready for you). ' +
    'Once those are filled in, run "Generate Advisor Roster" to build Attendance and MoM Log from them.'
  );
}

/**
 * Fully clears a sheet's data validation rules across its entire
 * grid, not just the currently-used range. sh.clear() does NOT
 * remove data validation (same gap as merges/banding) - without
 * this, a dropdown rule from an earlier column layout can survive a
 * rebuild and reject perfectly valid new values in that column.
 */
function clearAllValidations_(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
}

/**
 * Merges a range safely, even if a previous version of a rebuilt tab
 * left a differently-shaped merge behind. sh.clear() does not reliably
 * remove merged cells, so re-running a rebuild after the layout has
 * changed can otherwise throw "range and a merged cell partially
 * overlap" and silently halt the whole function midway through.
 */
function safeMerge_(sheet, a1Range) {
  const range = sheet.getRange(a1Range);
  try { range.breakApart(); } catch (e) {}
  return range.merge();
}

/**
 * Colors a title band WITHOUT merging it - use this instead of
 * safeMerge_ on any sheet that also freezes columns. A merge that
 * spans past the frozen-column boundary throws "can't freeze columns
 * which contain only part of a merged cell," so title bands on wide
 * tables (Attendance, MoM Log) use this instead.
 */
function bandTitle_(sheet, a1Range, text) {
  const range = sheet.getRange(a1Range);
  try { range.breakApart(); } catch (e) {}
  range.setBackground(COLOR_HEADER_BG);
  return range.getCell(1, 1).setValue(text)
    .setFontSize(15).setFontWeight('bold').setFontColor(COLOR_HEADER_TEXT);
}

// ---- SETTINGS SHEET ----
function setupSettingsSheet(ss) {
  let sh = ss.getSheetByName(SHEET_SETTINGS);
  if (sh) return; // don't overwrite an existing configured Settings tab
  sh = ss.insertSheet(SHEET_SETTINGS);

  const rows = [
    ['Field', 'Value', 'Active?'],
    ['AD Name', '', 'Active'],
    ['AD Email', '', 'Active'],
    ['HS Principal Name', '', 'Active'],
    ['HS Principal Email', '', 'Active'],
    ['HS Office Assistant Email', '', 'Not Active'],
    ['MS Principal Name', '', 'Not Active'],
    ['MS Principal Email', '', 'Not Active'],
    ['MS Office Assistant Email', '', 'Not Active'],
    ['Director Name', '', 'Not Active'],
    ['Director Email', '', 'Not Active'],
    ['DCI MS/HS Director Name', '', 'Not Active'],
    ['DCI MS/HS Director Email', '', 'Not Active'],
    ['ADCI MS/HS Director Email', '', 'Not Active'],
    ['SSC Manager Name', '', 'Not Active'],
    ['SSC Manager Email', '', 'Not Active'],
    ['', '', ''],
    ['Term # (1, 2, or 3)', 1, ''],
    ['Weeks per Trimester (e.g. 10)', 10, ''],
    ['Number of Advisors', 6, ''],
    ['Folder Term ID (advisor workbooks live here)', '', ''],
    ['Archive Term ID (CSV archive folder)', '', ''],
    ['MoM Folder ID (generated MoMs saved here)', '', ''],
    ['MoM Google Doc Template ID', '', ''],
    ['3-Week Report Google Doc Template ID', '', ''],
    ['Reports Folder ID (contains advisor subfolders)', '', ''],
    ['Advisor Checklist Google Doc Template ID', '', ''],
    ['Advisor Workbook Template Spreadsheet ID', '', ''],
    ['Academic Year (e.g. 26/27)', '26/27', ''],
    ['Email Footer - Sender Line (e.g. Office of the Academic Dean)', 'Office of the Academic Dean', ''],
    ['Email Footer - System Line (e.g. System American)', 'American System', ''],
    ['Email Footer - Campus Line (e.g. Main Campus)', 'Main Campus', ''],
    ['Email Footer - Tagline (e.g. Advisor Tracker)', 'Advisor Tracker', ''],
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground(COLOR_HEADER_BG).setFontColor(COLOR_HEADER_TEXT);

  // Active/Not Active dropdown on every role row (rows 2-16: AD through SSC Manager)
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Not Active'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, 3, 15, 1).setDataValidation(rule);
  addStatusColorRule(sh, sh.getRange(2, 3, 15, 1), { 'Active': STATUS_COMPLETED, 'Not Active': STATUS_NOT_STARTED });

  sh.getDataRange().setWrap(true);

  sh.setFrozenRows(1);
}

/** Reads a Settings field by its label in column A. */
function getSetting(label) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    throw new Error('No "Settings" tab found. Run "Advisor System > 1. Initialize Workbook" first.');
  }
  const data = sh.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === label) return data[i][1];
  }
  return '';
}

/** How many "Week N" tabs get created in each advisor workbook, and
 *  how many week columns Attendance/MoM Log have - set once in
 *  Settings ("Weeks per Trimester") instead of being fixed in code.
 *  Falls back to 10 if the setting is blank or invalid. */
function getWeekCount() {
  const n = Number(getSetting('Weeks per Trimester (e.g. 10)'));
  return (n && n > 0) ? Math.floor(n) : 10;
}

/** Returns array of {name, email} for a MoM's CC list, given the
 *  advisor's division ("MS" or "HS"). Division-specific roles (Principal,
 *  Office Assistant) only go out for that division; Director, DCI, ADCI
 *  and SSC Manager are division-agnostic and go out to everyone whenever
 *  they're marked Active - "Not Active" on any row means leave them out. */
function getActiveRecipients(division) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const map = {};
  data.forEach(r => map[String(r[0]).trim()] = r);

  const out = [];
  const addIfActive = (nameLabel, emailLabel) => {
    const emailRow = map[emailLabel];
    if (emailRow && emailRow[2] === 'Active' && emailRow[1]) {
      out.push({ name: map[nameLabel] ? map[nameLabel][1] : emailLabel, email: emailRow[1] });
    }
  };

  const div = String(division || '').trim().toUpperCase();
  if (div === 'HS') {
    addIfActive('HS Principal Name', 'HS Principal Email');
    addIfActive(null, 'HS Office Assistant Email');
  } else if (div === 'MS') {
    addIfActive('MS Principal Name', 'MS Principal Email');
    addIfActive(null, 'MS Office Assistant Email');
  }

  // Division-agnostic roles - go out to everyone if Active
  addIfActive('Director Name', 'Director Email');
  addIfActive('DCI MS/HS Director Name', 'DCI MS/HS Director Email');
  addIfActive(null, 'ADCI MS/HS Director Email');
  addIfActive('SSC Manager Name', 'SSC Manager Email');

  return out;
}