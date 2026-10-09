/**
 * ============================================================
 * ADVISOR WORKBOOK TEMPLATE SCRIPT
 * ============================================================
 * Paste this into Extensions > Apps Script on your ONE blank
 * "Advisor Workbook Template" Google Sheet - NOT on the main
 * Dean workbook, and not on individual advisor copies (they
 * inherit this automatically when the Dean's script copies the
 * template via Drive).
 *
 * Put that template file's ID into the main workbook's Settings
 * tab under "Advisor Workbook Template Spreadsheet ID".
 *
 * What this does:
 * - Watches the "Report - Week 3/6/9" tabs. Once every student's
 *   Trend/Key Concern/Recommendation is filled in, it waits a
 *   moment, asks the advisor to confirm, then finalizes: builds
 *   one report PDF per student from the Dean's report template,
 *   saves them in that advisor's Drive subfolder, writes each
 *   PDF's link back into the Report tab, and pushes the whole
 *   folder's link back to the Dean's Roster.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Advisor')
    .addItem('Enable Auto-Sync (run once)', 'enableAutoSync')
    .addItem('Generate 3-Week Report', 'generateReportMenuAction')
    .addToUi();
}

/** Menu item: generates the report immediately, without waiting for
 *  every row to be complete. Works out which week from whichever
 *  "Report - Week N" tab is currently open; if you're elsewhere, it asks. */
function generateReportMenuAction() {
  const ss = SpreadsheetApp.getActive();
  const active = ss.getActiveSheet();
  const match = active.getName().match(/^Report - Week (3|6|9)$/);
  let week;
  if (match) {
    week = parseInt(match[1], 10);
  } else {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.prompt('Generate 3-Week Report', 'Which week - 3, 6, or 9?', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    week = parseInt(resp.getResponseText().trim(), 10);
    if ([3, 6, 9].indexOf(week) === -1) {
      ui.alert('Enter 3, 6, or 9.');
      return;
    }
  }
  finalizeReport(week);
}

/** Advisor (or the Dean) runs this once per new workbook copy. */
function enableAutoSync() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'advisorWorkbookOnEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('advisorWorkbookOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert(
    'Auto-sync enabled. Once you finish filling in a Week 3, 6 or 9 report for every ' +
    'student, you\'ll be asked to finalize it and the reports will generate automatically.'
  );
}

/** Installable onEdit trigger (has full authorization, unlike a
 *  simple onEdit, so it's allowed to open the main workbook by ID
 *  and prompt with a dialog). */
function advisorWorkbookOnEdit(e) {
  try {
    const sh = e.range.getSheet();
    const match = sh.getName().match(/^Report - Week (3|6|9)$/);
    if (!match) return;

    const week = parseInt(match[1], 10);
    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (row < 6 || col < 2 || col > 4) return; // only Trend/Concern/Recommendation edits matter

    if (sh.getRange('B4').getValue() === 'Finalized') return; // already done - don't re-prompt
    if (!isReportComplete_(sh)) return;

    Utilities.sleep(3000); // give them a beat before interrupting
    if (sh.getRange('B4').getValue() === 'Finalized') return; // could have finished in the meantime

    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      'Finalize Report',
      'All students are filled in for this report. Finalize and generate the reports now?',
      ui.ButtonSet.YES_NO
    );
    if (resp === ui.Button.YES) finalizeReport(week);
  } catch (err) {
    // never let a trigger error block the advisor's editing
  }
}

function isReportComplete_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return false;
  const data = sheet.getRange(6, 1, lastRow - 5, 4).getValues();
  return data.length > 0 && data.every(r => r[0] && r[1] && r[2] && r[3]);
}

/**
 * Builds one report PDF per student (from the Dean's report
 * template), saves them in a "Week N" subfolder inside this
 * advisor's Drive folder (inside the Reports Folder), writes each
 * PDF link back onto this Report tab, and pushes that Week N
 * subfolder's link back to the Dean's Week N tab and Roster.
 */
function finalizeReport(week) {
  const ss = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName('Config');
  const mainId = cfg.getRange('B1').getValue();
  const advisorName = cfg.getRange('B2').getValue();
  const mainSs = SpreadsheetApp.openById(mainId);

  const templateId = getMainSetting_(mainSs, '3-Week Report Google Doc Template ID');
  const reportsFolderId = getMainSetting_(mainSs, 'Reports Folder ID (contains advisor subfolders)');
  const term = getMainSetting_(mainSs, 'Term # (1, 2, or 3)') || 1;
  if (!templateId || !reportsFolderId) {
    SpreadsheetApp.getUi().alert(
      'Ask the Academic Dean to set the "3-Week Report Google Doc Template ID" and ' +
      '"Reports Folder ID" in Settings first - the reports can\'t be generated without those.'
    );
    return;
  }

  const reportsParent = DriveApp.getFolderById(reportsFolderId);
  const advisorFolder = getOrCreateAdvisorFolder_(reportsParent, advisorName);
  const weekFolder = getOrCreateSubfolder_(advisorFolder, 'Week ' + week);

  const tabName = 'Report - Week ' + week;
  const sh = ss.getSheetByName(tabName);
  const numRows = sh.getLastRow() - 5;
  const data = sh.getRange(6, 1, numRows, 4).getValues(); // Name, Trend, Concern, Recommendation
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy');

  for (let i = 0; i < data.length; i++) {
    const [name, trend, concern, recommendation] = data[i];
    if (!name) continue;

    const baseName = name + '_T' + term + 'W' + week + '_' + advisorName;
    const docFile = DriveApp.getFileById(templateId).makeCopy(baseName, weekFolder);
    const doc = DocumentApp.openById(docFile.getId());
    const body = doc.getBody();
    body.replaceText('{{STUDENT_NAME}}', name);
    body.replaceText('{{ADVISOR_NAME}}', advisorName);
    body.replaceText('{{TERM}}', String(term));
    body.replaceText('{{WEEK}}', String(week));
    body.replaceText('{{TREND}}', trend || '');
    body.replaceText('{{KEY_CONCERN}}', concern || '');
    body.replaceText('{{RECOMMENDATION}}', recommendation || '');
    body.replaceText('{{DATE}}', dateStr);
    doc.saveAndClose();

    const pdfBlob = docFile.getAs('application/pdf');
    const pdfFile = weekFolder.createFile(pdfBlob).setName(baseName + '.pdf');
    try {
      pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // org sharing policy may block this - the file still exists and is
      // visible to the Dean/advisor as owner/editor, just not link-shared
    }
    docFile.setTrashed(true); // was only a working copy - the PDF is the deliverable

    sh.getRange(6 + i, 5).setValue(pdfFile.getUrl());
  }

  sh.getRange('B4').setValue('Finalized');

  const folderUrl = weekFolder.getUrl();
  const label = 'Week ' + week + ' Reports Folder Link';
  const weekSheet = ss.getSheetByName('Week ' + week);
  if (weekSheet) setLabelValue_(weekSheet, label, folderUrl);

  const roster = mainSs.getSheetByName('Advisor Roster');
  if (roster) {
    const rdata = roster.getDataRange().getValues();
    const header = rdata[1]; // row 1 = title band, row 2 = column headers
    const colIdx = header.indexOf('Week ' + week + ' Report Link');
    if (colIdx !== -1) {
      for (let r = 2; r < rdata.length; r++) { // data starts row 3
        if (rdata[r][0] === advisorName) {
          roster.getRange(r + 1, colIdx + 1).setValue(folderUrl);
          break;
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    'Done! Generated ' + data.filter(r => r[0]).length + ' report PDFs in your "' +
    advisorName + ' / Week ' + week + '" folder. The folder link has been sent to the Academic Dean automatically.'
  );
}

function getMainSetting_(mainSs, label) {
  const sh = mainSs.getSheetByName('Settings');
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === label) return data[i][1];
  }
  return '';
}

/** Finds (or creates) the subfolder named exactly after this advisor,
 *  inside the Dean's Reports Folder for this term. */
function getOrCreateAdvisorFolder_(parentFolder, advisorName) {
  const it = parentFolder.getFoldersByName(advisorName);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(advisorName);
}

/** Finds (or creates) a named subfolder inside a given folder -
 *  used for the "Week 3" / "Week 6" / "Week 9" subfolders inside
 *  each advisor's own folder. */
function getOrCreateSubfolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

function setLabelValue_(sheet, label, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  const data = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).indexOf(label) === 0) {
      sheet.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
}