const NOTES_SHEET = 'Notes';
const EDITS_SHEET = 'Edits';
const STATUS_SHEET = 'Status';

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || 'callback';
  let data;
  if (params.kind === 'notes') {
    data = listNotes_();
  } else if (params.kind === 'edit') {
    data = getEdit_(params.note_id);
  } else {
    data = listStatus_();
  }
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  const body = JSON.parse((e.postData && e.postData.contents) || '{}');
  if (body.kind === 'note') {
    upsertNote_(body);
  } else if (body.kind === 'edit') {
    upsertEdit_(body);
  } else {
    upsertStatus_(body);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name, headers) {
  const ss = ss_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  const missing = headers.filter(header => !current.includes(header));
  if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  return sheet;
}

function noteHeaders_() {
  return ['note_id', 'date', 'title', 'subtitle', 'method', 'product', 'keywords', 'status', 'note_type', 'archived', 'updated_at'];
}

function editHeaders_() {
  return ['note_id', 'note_type', 'payload', 'updated_at'];
}

function statusHeaders_() {
  return ['note_id', 'status', 'finished_date', 'outcome', 'comment', 'updated_at'];
}

function rowsAsObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => obj[header] = row[index]);
    return obj;
  });
}

function rowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return -1;
  const headers = values[0].map(String);
  const idCol = headers.indexOf('note_id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) return i + 1;
  }
  return -1;
}

function writeObject_(sheet, headers, rowNumber, object) {
  const row = headers.map(header => object[header] == null ? '' : object[header]);
  if (rowNumber > 0) {
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function upsertNote_(body) {
  if (!body.note_id) throw new Error('note_id is required');
  const headers = noteHeaders_();
  const sheet = sheet_(NOTES_SHEET, headers);
  const existingRow = rowById_(sheet, body.note_id);
  const existing = existingRow > 0 ? rowsAsObjects_(sheet)[existingRow - 2] : {};
  const record = {
    note_id: body.note_id,
    date: body.date || existing.date || '',
    title: body.title || existing.title || '',
    subtitle: body.subtitle || existing.subtitle || '',
    method: body.method || existing.method || '',
    product: body.product || existing.product || '',
    keywords: body.keywords || existing.keywords || '',
    status: body.status || existing.status || 'ongoing',
    note_type: body.note_type || existing.note_type || 'synthesis',
    archived: String(Boolean(body.archived)),
    updated_at: body.updated_at || new Date().toISOString()
  };
  writeObject_(sheet, headers, existingRow, record);
}

function upsertEdit_(body) {
  if (!body.note_id) throw new Error('note_id is required');
  const headers = editHeaders_();
  const sheet = sheet_(EDITS_SHEET, headers);
  const existingRow = rowById_(sheet, body.note_id);
  const record = {
    note_id: body.note_id,
    note_type: body.note_type || (body.payload && body.payload.note_type) || 'synthesis',
    payload: JSON.stringify(body.payload || {}),
    updated_at: body.updated_at || new Date().toISOString()
  };
  writeObject_(sheet, headers, existingRow, record);
}

function upsertStatus_(body) {
  if (!body.note_id) throw new Error('note_id is required');
  const headers = statusHeaders_();
  const sheet = sheet_(STATUS_SHEET, headers);
  const existingRow = rowById_(sheet, body.note_id);
  const record = {
    note_id: body.note_id,
    status: body.status || 'ongoing',
    finished_date: body.finished_date || '',
    outcome: body.outcome || '',
    comment: body.comment || '',
    updated_at: body.updated_at || new Date().toISOString()
  };
  writeObject_(sheet, headers, existingRow, record);
}

function listNotes_() {
  const sheet = sheet_(NOTES_SHEET, noteHeaders_());
  return rowsAsObjects_(sheet)
    .filter(row => String(row.archived).toLowerCase() !== 'true')
    .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
}

function getEdit_(noteId) {
  if (!noteId) return {};
  const sheet = sheet_(EDITS_SHEET, editHeaders_());
  const row = rowsAsObjects_(sheet).find(item => String(item.note_id) === String(noteId));
  if (!row) return {};
  let payload = {};
  try {
    payload = JSON.parse(row.payload || '{}');
  } catch (error) {
    payload = {};
  }
  return { note_id: row.note_id, note_type: row.note_type || payload.note_type || 'synthesis', payload, updated_at: row.updated_at };
}

function listStatus_() {
  const sheet = sheet_(STATUS_SHEET, statusHeaders_());
  return rowsAsObjects_(sheet);
}
