/**
 * Google Apps Script backend for the AI Native Working Practice Survey.
 *
 * Deploy as a Web app, then paste the /exec URL into APPS_SCRIPT_URL
 * in index.html. The first submission creates the sheet headers.
 */
function joinValues(value) {
  return Array.isArray(value) ? value.join(' | ') : (value || '');
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents || '{}');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Name', 'Email', 'Role', 'Role (other)', 'Frequency', 'Highest Level', 'Persistent Context',
        'AI Actions', 'Assurance', 'Reuse', 'Most Complex AI Task', 'Limits',
        'Session Topics', 'Questions for the Session', 'Duration (seconds)'
      ]);
      sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name || '',
      data.email || '',
      joinValues(data.role),
      data.role_other || '',
      data.frequency || '',
      data.level || '',
      joinValues(data.context),
      joinValues(data.ai_actions),
      joinValues(data.assurance),
      data.reuse || '',
      data.complex_task || '',
      joinValues(data.limits),
      joinValues(data.session_topics),
      data.questions || '',
      data.duration_seconds || ''
    ]);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'AI Native survey backend is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
