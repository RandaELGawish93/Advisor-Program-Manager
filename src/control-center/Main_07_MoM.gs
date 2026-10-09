/**
 * MoM generation and sending. Your Google Doc template needs:
 *
 * 1. Plain-text tags anywhere in the header area:
 *    {{DATE}} {{TIME}} {{LOCATION}} {{ADVISOR_NAME}} {{DIVISION}}
 *    {{GENERAL_NOTES}} {{FOLLOWUP_ACTIONS}} {{NEXT_MEETING_DATE}}
 *
 * 2. Your real "Students Discussed" TABLE (Student Name | Status
 *    Update | Concerns | Next Steps), with its one data row's
 *    cells set to: {{STUDENT_NAME}} | {{STATUS_UPDATE}} | {{CONCERNS}}
 *    | {{NEXT_STEPS}}  - the script duplicates that row once per
 *    student and fills each column, then removes the placeholder.
 *
 * Flow: generate the Doc (saved to your MoM Folder) -> preview subject/
 * To/Cc/body in a popup, where you can also type your own feedback note
 * on the advisor's week -> you click Send -> your note is written into
 * the Doc's General Notes and the email, the PDF is created from the
 * finalized Doc, and the email goes out with that PDF attached. The Doc
 * stays on Drive as the editable original.
 */

/** Menu item: run with the cursor on the advisor's row in EITHER
 *  the Attendance sheet or the MoM Log sheet; asks which week, then
 *  opens the preview popup (nothing is sent until you click Send there). */
function sendMoMForSelectedRow() {
  const ss = SpreadsheetApp.getActive();
  const activeSheet = ss.getActiveSheet();
  const activeCell = ss.getActiveCell();
  const advisorName = activeSheet.getRange(activeCell.getRow(), 1).getValue();

  if (!advisorName || !getAdvisorList().some(a => a.name === advisorName)) {
    SpreadsheetApp.getUi().alert('Click a cell on an advisor\'s row in the Attendance or MoM Log tab first.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Send MoM', 'Week number (1-10) for ' + advisorName + ':', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const week = parseInt(resp.getResponseText().trim(), 10);
  if (!week || week < 1 || week > 10) {
    ui.alert('Enter a valid week number 1-10.');
    return;
  }

  buildMoMPreview(advisorName, week);
}

/** Builds the MoM Doc (leaving {{GENERAL_NOTES}} untouched for now),
 *  then opens a preview popup showing Subject / To / Cc / body, plus
 *  a field where you type your own feedback on the advisor's week -
 *  what they did well, what to adjust. Nothing is sent, and the PDF
 *  isn't created, until the "Send Email" button is clicked; your
 *  note gets written into the Doc's General Notes section and into
 *  the email itself at that point. */
function buildMoMPreview(advisorName, week) {
  const templateId = getSetting('MoM Google Doc Template ID');
  const momFolderId = getSetting('MoM Folder ID (generated MoMs saved here)');
  if (!templateId || !momFolderId) {
    throw new Error('Set "MoM Google Doc Template ID" and "MoM Folder ID" in Settings first.');
  }
  const roster = getRosterRow(advisorName);
  const division = roster ? roster.values[2] : '';
  const advisorEmail = roster ? roster.values[1] : '';
  const followUpDay = roster ? roster.values[4] : '';
  if (!advisorEmail) {
    SpreadsheetApp.getUi().alert('No email on file for ' + advisorName + ' - add it in Advisor Roster first.');
    return;
  }

  const term = getSetting('Term # (1, 2, or 3)') || 1;
  const academicYear = String(getSetting('Academic Year (e.g. 26/27)') || '').trim();
  const studentRows = getStudentDataForWeek(advisorName, week);
  const nextMeetingDate = getNextMeetingDateStr(followUpDay);

  const momFolder = DriveApp.getFolderById(momFolderId);
  const baseName = 'MoM_T' + term + 'W' + week + '_' + advisorName;
  const docFile = DriveApp.getFileById(templateId).makeCopy(baseName, momFolder);
  const doc = DocumentApp.openById(docFile.getId());
  const body = doc.getBody();

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy');
  body.replaceText('{{DATE}}', dateStr);
  body.replaceText('{{TIME}}', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'));
  body.replaceText('{{LOCATION}}', "Academic Dean's Office");
  body.replaceText('{{ADVISOR_NAME}}', advisorName);
  body.replaceText('{{DIVISION}}', division);
  body.replaceText('{{FOLLOWUP_ACTIONS}}', '');
  body.replaceText('{{NEXT_MEETING_DATE}}', nextMeetingDate);
  fillStudentsTable(doc, studentRows);
  doc.saveAndClose();
  // {{GENERAL_NOTES}} is deliberately left in place - filled in confirmSendMoM
  // once you've typed your feedback in the preview popup.

  const subject = (academicYear ? academicYear + ': ' : '') +
    'Advisor MoM T' + term + 'W' + week + ' - ' + advisorName;
  const ccList = getActiveRecipients(division);
  const ccEmails = ccList.map(r => r.email).filter(Boolean);
  const htmlBody = buildMoMEmailHtml(advisorName, week, term, academicYear, dateStr);

  showMoMPreviewDialog({
    advisorName: advisorName,
    week: week,
    term: term,
    docFileId: docFile.getId(),
    momFolderId: momFolderId,
    toEmail: advisorEmail,
    ccEmails: ccEmails,
    ccNames: ccList.map(r => r.name).filter(Boolean),
    subject: subject,
    htmlBody: htmlBody,
  });
}

/** Renders the Subject/To/Cc/body preview popup with Send/Cancel,
 *  plus a feedback box for the Dean's own note on the advisor's week. */
function showMoMPreviewDialog(data) {
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:13px;color:${COLOR_READONLY_TXT};padding:4px 8px;">
      <div style="margin-bottom:10px;">
        <label style="font-weight:bold;display:block;margin-bottom:2px;">Subject</label>
        <input id="subject" type="text" value="${escapeHtml_(data.subject)}"
          style="width:100%;padding:6px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
      </div>
      <div style="margin-bottom:10px;">
        <label style="font-weight:bold;display:block;margin-bottom:2px;">To</label>
        <div style="padding:6px;background:${COLOR_READONLY_BG};border-radius:4px;">${escapeHtml_(data.toEmail)}</div>
      </div>
      <div style="margin-bottom:14px;">
        <label style="font-weight:bold;display:block;margin-bottom:2px;">Cc (edit if needed)</label>
        <input id="cc" type="text" value="${escapeHtml_(data.ccEmails.join(', '))}"
          style="width:100%;padding:6px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
        <div style="font-size:11px;color:#888;margin-top:2px;">${escapeHtml_(data.ccNames.join(', ') || 'No one else active for this division.')}</div>
      </div>
      <div style="margin-bottom:14px;">
        <label style="font-weight:bold;display:block;margin-bottom:2px;">Your Note for ${escapeHtml_(data.advisorName)} (optional)</label>
        <div style="font-size:11px;color:#888;margin-bottom:4px;">What they did well this week, what to adjust, any feedback - goes into the MoM's General Notes and the email below.</div>
        <textarea id="deanNote" rows="4" placeholder="e.g. Great job following up on missing homework - let's also make sure the weekly status line always names a specific reason, not just 'Better/Worse'."
          style="width:100%;padding:6px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;font-family:inherit;font-size:13px;"
          oninput="updateNotePreview(this.value)"></textarea>
      </div>
      <label style="font-weight:bold;display:block;margin-bottom:4px;">Email preview</label>
      <div id="emailPreview" style="border:1px solid #ddd;border-radius:6px;overflow:hidden;margin-bottom:14px;max-height:280px;overflow-y:auto;">
        ${data.htmlBody}
      </div>
      <div style="text-align:right;">
        <button onclick="google.script.host.close()"
          style="padding:8px 16px;border:1px solid #ccc;border-radius:4px;background:white;cursor:pointer;margin-right:8px;">Cancel</button>
        <button id="sendBtn" onclick="send()"
          style="padding:8px 20px;border:none;border-radius:4px;background:${COLOR_HEADER_BG};color:white;font-weight:bold;cursor:pointer;">Send Email</button>
      </div>
    </div>
    <script>
      function updateNotePreview(val) {
        var box = document.getElementById('dean-note-box');
        var txt = document.getElementById('dean-note-text');
        if (!box || !txt) return;
        if (val && val.trim()) {
          txt.innerText = val;
          box.style.display = 'block';
        } else {
          box.style.display = 'none';
        }
      }
      function send() {
        document.getElementById('sendBtn').disabled = true;
        document.getElementById('sendBtn').innerText = 'Sending...';
        google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
          .withFailureHandler(function(err){ alert('Error: ' + err.message); document.getElementById('sendBtn').disabled=false; })
          .confirmSendMoM({
            advisorName: ${JSON.stringify(data.advisorName)},
            week: ${data.week},
            term: ${data.term},
            docFileId: ${JSON.stringify(data.docFileId)},
            momFolderId: ${JSON.stringify(data.momFolderId)},
            toEmail: ${JSON.stringify(data.toEmail)},
            ccEmailsStr: document.getElementById('cc').value,
            subject: document.getElementById('subject').value,
            htmlBody: ${JSON.stringify(data.htmlBody)},
            deanNote: document.getElementById('deanNote').value
          });
      }
    </script>
  `;
  const output = HtmlService.createHtmlOutput(html).setWidth(560).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(output, 'Preview MoM Email');
}

/** Called from the preview popup's Send button. Fills in the Dean's
 *  note, builds the PDF from the finalized Doc, sends the email,
 *  logs it, and flips the Dean-controlled fields to Done. */
function confirmSendMoM(payload) {
  const note = String(payload.deanNote || '').trim();

  const docFile = DriveApp.getFileById(payload.docFileId);
  const doc = DocumentApp.openById(payload.docFileId);
  doc.getBody().replaceText('{{GENERAL_NOTES}}', note);
  doc.saveAndClose();

  const momFolder = DriveApp.getFolderById(payload.momFolderId);
  const pdfFile = momFolder.createFile(docFile.getAs('application/pdf')).setName(docFile.getName() + '.pdf');
  try { pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const pdfBlob = pdfFile.getBlob();

  let finalHtml = payload.htmlBody;
  if (note) {
    finalHtml = finalHtml
      .replace('id="dean-note-box" style="display:none;', 'id="dean-note-box" style="display:block;')
      .replace('<div id="dean-note-text" style="margin-top:4px;"></div>', '<div id="dean-note-text" style="margin-top:4px;">' + escapeHtml_(note).replace(/\n/g, '<br>') + '</div>');
  }

  const ccEmailsStr = String(payload.ccEmailsStr || '').trim();
  MailApp.sendEmail({
    to: payload.toEmail,
    cc: ccEmailsStr,
    subject: payload.subject,
    htmlBody: finalHtml,
    attachments: [pdfBlob],
  });

  const ccEmails = ccEmailsStr ? ccEmailsStr.split(',').map(s => s.trim()).filter(Boolean) : [];
  logEmailSent(payload.advisorName, payload.week, pdfFile.getUrl(), ccEmails);
  writeMoMLinkToLog(payload.advisorName, payload.week, pdfFile.getUrl(),
    'MoM T' + payload.term + 'W' + payload.week + ' (PDF version)');
  markAttendanceCompleted(payload.advisorName, payload.week);
  pushAttendDone(payload.advisorName, payload.week);
  pushMoMReceivedDone(payload.advisorName, payload.week);

  return 'sent';
}

/** Modern, minimal HTML email body using your Sheets theme colors,
 *  with a footer pulled from Settings so you can update it without
 *  touching code. */
function buildMoMEmailHtml(advisorName, week, term, academicYear, dateStr) {
  const adLine = getSetting('Email Footer - Sender Line (e.g. Office of the Academic Dean)') || 'Office of the Academic Dean';
  const sysLine = getSetting('Email Footer - System Line (e.g. System American)') || 'System American';
  const campusLine = getSetting('Email Footer - Campus Line (e.g. Main Campus)') || 'Main Campus';
  const tagline = getSetting('Email Footer - Tagline (e.g. Advisor Tracker)') || 'Advisor Tracker';
  const yearLabel = academicYear ? ('AY ' + academicYear) : '';

  return `
  <div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto;border:1px solid #e6e6e6;border-radius:8px;overflow:hidden;">
    <div style="background:${COLOR_HEADER_BG};padding:18px 24px;">
      <div style="color:#ffffff;font-size:17px;font-weight:bold;">Advisor - Minutes of Meeting</div>
      <div style="color:#cfe2f3;font-size:12px;margin-top:2px;">Term ${term} · Week ${week}${yearLabel ? ' · ' + yearLabel : ''}</div>
    </div>
    <div style="padding:22px 24px;color:${COLOR_READONLY_TXT};font-size:13px;line-height:1.6;">
      <p style="margin-top:0;">Dear ${escapeHtml_(advisorName)},</p>
      <p>Attached is the Minutes of Meeting (PDF) for your Week ${week} Advisor check-in, held on ${dateStr}.</p>
      <div id="dean-note-box" style="display:none;background:${COLOR_KPI_BG};color:${COLOR_KPI_TXT};border-radius:6px;padding:10px 14px;margin:16px 0;font-size:12.5px;">
        <b>Note from the Academic Dean:</b>
        <div id="dean-note-text" style="margin-top:4px;"></div>
      </div>
      <div style="background:${COLOR_SECTION_BG};color:${COLOR_SECTION_TXT};border-radius:6px;padding:10px 14px;margin:16px 0;font-size:12px;">
        📄 Please keep this for your records. If anything here needs correction, reply to this email.
      </div>
      <p style="margin-bottom:0;">Thank you for the work you put into supporting your students this week.</p>
    </div>
    <div style="background:${COLOR_READONLY_BG};padding:14px 24px;font-size:11px;color:#6b6b6b;line-height:1.6;border-top:1px solid #e6e6e6;">
      <div>${escapeHtml_(adLine)}</div>
      <div>${escapeHtml_(sysLine)}</div>
      <div>${escapeHtml_(campusLine)}</div>
      <div style="margin-top:4px;font-style:italic;">${escapeHtml_(tagline)}</div>
    </div>
  </div>`;
}

function escapeHtml_(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Computes the advisor's actual next meeting date: the next real
 * calendar occurrence of their Weekly Follow-Up Day, counting from
 * TODAY - not from whichever week's MoM you happen to be sending.
 * That matters when you're catching up on an older week's MoM after
 * the fact: tying "next meeting" to that week + 1 could land on a
 * date already in the past. This always looks forward from today,
 * so it's never wrong regardless of which week you're sending for.
 * If today itself is the follow-up day, "next" means next week's
 * occurrence, not today.
 */
function getNextMeetingDateStr(followUpDayName) {
  const dayOffsets = {
    sunday: 0, monday: 1, tuesday: 2, wednesday: 3, thursday: 4, friday: 5, saturday: 6,
  };
  const offset = dayOffsets[String(followUpDayName || '').trim().toLowerCase()];
  if (offset === undefined) return '';

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let daysUntil = (offset - today.getDay() + 7) % 7;
  if (daysUntil === 0) daysUntil = 7; // today IS the follow-up day - next one is next week

  const next = new Date(today);
  next.setDate(next.getDate() + daysUntil);
  return Utilities.formatDate(next, Session.getScriptTimeZone(), 'dd MMMM yyyy');
}

/**
 * Pulls each student's Weekly Status / Why / Notes from the
 * advisor's own Week N tab if that workbook can be reached;
 * otherwise falls back to just the student names with blank fields.
 * Column layout in the Week N student table (0-indexed):
 * 0 ID | 1 Name | 2 Class | 3 Subjects | 4-11 FC..Attendance |
 * 12 Weekly Status | 13 Why | 14 Flag Counsellor | 15 On Track | 16 Notes
 */
function getStudentDataForWeek(advisorName, week) {
  const roster = getStudentList().filter(s => s[3] === advisorName);
  const fallback = roster.map(s => ({ name: s[1], status: '', concerns: '', nextSteps: '' }));

  const folderId = getSetting('Folder Term ID (advisor workbooks live here)');
  if (!folderId) return fallback;
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName('Advisor Workbook - ' + advisorName);
    if (!files.hasNext()) return fallback;
    const childSs = SpreadsheetApp.openById(files.next().getId());
    const weekSheet = childSs.getSheetByName('Week ' + week);
    if (!weekSheet) return fallback;

    const data = weekSheet.getDataRange().getValues();
    const headerRowIdx = data.findIndex(r => r[0] === 'Student ID');
    if (headerRowIdx === -1) return fallback;

    const out = [];
    for (let r = headerRowIdx + 1; r < data.length; r++) {
      const name = data[r][1];
      if (!name) break; // end of student block
      const weeklyStatus = data[r][12] || '';
      const why = data[r][13] || '';
      const flagCounsellor = data[r][14] === true ? 'Counsellor check-in flagged' : '';
      const onTrack = data[r][15] || '';
      const notes = data[r][16] || '';
      out.push({
        name: name,
        status: [weeklyStatus, why].filter(Boolean).join(' - '),
        concerns: [onTrack, flagCounsellor, notes].filter(Boolean).join(' - '),
        nextSteps: '',
      });
    }
    return out.length ? out : fallback;
  } catch (err) {
    return fallback; // advisor workbook not reachable yet - use names only
  }
}

/**
 * Finds the table whose data row contains {{STUDENT_NAME}}, fills
 * the first student into that row, then inserts one new row per
 * remaining student right after it with the same 4 columns.
 */
function fillStudentsTable(doc, students) {
  const body = doc.getBody();
  let table = null;
  let rowIndex = -1;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    const t = child.asTable();
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      if (row.getNumCells() > 0 && row.getCell(0).getText().indexOf('{{STUDENT_NAME}}') !== -1) {
        table = t;
        rowIndex = r;
        break;
      }
    }
    if (table) break;
  }
  if (!table) return; // template has no tagged table row - nothing to fill

  const numCols = table.getRow(rowIndex).getNumCells();
  const setRow = (row, s) => {
    row.getCell(0).setText(s.name || '');
    if (numCols > 1) row.getCell(1).setText(s.status || '');
    if (numCols > 2) row.getCell(2).setText(s.concerns || '');
    if (numCols > 3) row.getCell(3).setText(s.nextSteps || '');
  };

  if (students.length === 0) {
    setRow(table.getRow(rowIndex), { name: '', status: '', concerns: '', nextSteps: '' });
    return;
  }

  setRow(table.getRow(rowIndex), students[0]);
  for (let i = 1; i < students.length; i++) {
    const newRow = table.insertTableRow(rowIndex + i);
    for (let c = 0; c < numCols; c++) newRow.appendTableCell('');
    setRow(newRow, students[i]);
  }
}

/** Writes a labeled hyperlink (e.g. "MoM T1W3 (PDF version)") into
 *  the MoM Log's Week N cell for this advisor. */
function writeMoMLinkToLog(advisorName, week, url, label) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MOM_LOG);
  const data = sh.getDataRange().getValues();
  const header = data[ROSTER_HEADER_ROW - 1];
  const colIdx = header.indexOf('Week ' + week + ' MoM Link');
  for (let r = ROSTER_HEADER_ROW; r < data.length; r++) {
    if (data[r][0] === advisorName) {
      sh.getRange(r + 1, colIdx + 1).setFormula('=HYPERLINK("' + url + '","' + (label || url) + '")');
      return;
    }
  }
}

function markAttendanceCompleted(advisorName, week) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ATTENDANCE);
  const data = sh.getDataRange().getValues();
  const header = data[ROSTER_HEADER_ROW - 1];
  const colIdx = header.indexOf('Week ' + week);
  for (let r = ROSTER_HEADER_ROW; r < data.length; r++) {
    if (data[r][0] === advisorName) {
      sh.getRange(r + 1, colIdx + 1).setValue('Completed');
      return;
    }
  }
}

/** Flips "Attend Weekly Meeting" to Done on the profile tab and the
 *  advisor's own Week N tab. Called both when you mark Attendance
 *  Completed directly, and when you send the MoM. */
function pushAttendDone(advisorName, week) {
  pushLabeledDone_(advisorName, week, 'Attend Weekly Meeting');
}

/** Flips "MoM Received" to Done - only called once the MoM email
 *  actually goes out. */
function pushMoMReceivedDone(advisorName, week) {
  pushLabeledDone_(advisorName, week, 'MoM Received');
}

function pushLabeledDone_(advisorName, week, label) {
  const ss = SpreadsheetApp.getActive();
  const mainProfile = ss.getSheetByName(PROFILE_PREFIX + advisorName);
  if (mainProfile) setLabeledCell(mainProfile, label, 'Done');

  const folderId = getSetting('Folder Term ID (advisor workbooks live here)');
  if (!folderId) return;
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName('Advisor Workbook - ' + advisorName);
    if (!files.hasNext()) return;
    const childSs = SpreadsheetApp.openById(files.next().getId());
    const weekSheet = childSs.getSheetByName('Week ' + week);
    if (!weekSheet) return;
    setLabeledCell(weekSheet, label, 'Done');
  } catch (err) {
    // advisor workbook not reachable yet - safe to ignore
  }
}