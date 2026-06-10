const APP_NAME = 'Experimental Notes Sheets API';
const APP_VERSION = '2026-06-10-v8-note-type-repair';
const APP_UPDATED_AT = '2026-06-10T19:55:00+09:00';
const NOTES_SHEET = 'Notes';
const EDITS_SHEET = 'Edits';
const STATUS_SHEET = 'Status';

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || '';
  let data;
  if (params.kind === 'version' || (!params.kind && !callback)) {
    data = version_();
  } else if (params.kind === 'notes') {
    data = listNotes_(params.include_archived === '1' || params.include_archived === 'true');
  } else if (params.kind === 'edit') {
    data = getEdit_(params.note_id);
  } else {
    data = listStatus_();
  }
  if (!callback) {
    return ContentService
      .createTextOutput(JSON.stringify(data, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
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
    .createTextOutput(JSON.stringify({ ok: true, app_name: APP_NAME, app_version: APP_VERSION }))
    .setMimeType(ContentService.MimeType.JSON);
}

function version_() {
  return {
    ok: true,
    app_name: APP_NAME,
    app_version: APP_VERSION,
    updated_at: APP_UPDATED_AT,
    sheets: {
      notes: NOTES_SHEET,
      edits: EDITS_SHEET,
      status: STATUS_SHEET
    }
  };
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
    headers.forEach((header, index) => obj[header] = normalizeCell_(row[index], header));
    return obj;
  });
}

function normalizeCell_(value, header) {
  if (value instanceof Date) {
    const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
    if (header === 'date' || header === 'finished_date') {
      return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
    }
    return Utilities.formatDate(value, timezone, "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  return value;
}

function rowsById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  const idCol = headers.indexOf('note_id');
  const updatedCol = headers.indexOf('updated_at');
  const matches = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      matches.push({ rowNumber: i + 1, updated_at: updatedCol >= 0 ? normalizeCell_(values[i][updatedCol], 'updated_at') : '' });
    }
  }
  return matches.sort((a, b) => String(b.updated_at || '').localeCompare(String(a.updated_at || '')));
}

function rowById_(sheet, id) {
  const rows = rowsById_(sheet, id);
  return rows.length ? rows[0].rowNumber : -1;
}

function deleteDuplicateRowsById_(sheet, id, keepRowNumber) {
  rowsById_(sheet, id)
    .map(item => item.rowNumber)
    .filter(rowNumber => rowNumber !== keepRowNumber)
    .sort((a, b) => b - a)
    .forEach(rowNumber => sheet.deleteRow(rowNumber));
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
  const editTypeById = latestEditTypes_();
  const record = {
    note_id: body.note_id,
    date: body.date || existing.date || '',
    title: body.title || existing.title || '',
    subtitle: body.subtitle || existing.subtitle || '',
    method: body.method || existing.method || '',
    product: body.product || existing.product || '',
    keywords: body.keywords || existing.keywords || '',
    status: body.status || existing.status || 'ongoing',
    note_type: body.note_type || existing.note_type || editTypeById[body.note_id] || '',
    archived: String(Boolean(body.archived)),
    updated_at: body.updated_at || new Date().toISOString()
  };
  writeObject_(sheet, headers, existingRow, record);
  deleteDuplicateRowsById_(sheet, body.note_id, existingRow > 0 ? existingRow : sheet.getLastRow());
}

function upsertEdit_(body) {
  if (!body.note_id) throw new Error('note_id is required');
  const headers = editHeaders_();
  const sheet = sheet_(EDITS_SHEET, headers);
  const existingRow = rowById_(sheet, body.note_id);
  const existing = existingRow > 0 ? rowsAsObjects_(sheet)[existingRow - 2] : {};
  const record = {
    note_id: body.note_id,
    note_type: body.note_type || (body.payload && body.payload.note_type) || existing.note_type || '',
    payload: JSON.stringify(body.payload || {}),
    updated_at: body.updated_at || new Date().toISOString()
  };
  writeObject_(sheet, headers, existingRow, record);
  deleteDuplicateRowsById_(sheet, body.note_id, existingRow > 0 ? existingRow : sheet.getLastRow());
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
  deleteDuplicateRowsById_(sheet, body.note_id, existingRow > 0 ? existingRow : sheet.getLastRow());
}

function listNotes_(includeArchived) {
  const sheet = sheet_(NOTES_SHEET, noteHeaders_());
  const editTypeById = latestEditTypes_();
  const latest = {};
  rowsAsObjects_(sheet).forEach(row => {
    if (!row.note_id) return;
    if (!row.note_type && editTypeById[row.note_id]) row.note_type = editTypeById[row.note_id];
    const previous = latest[row.note_id];
    if (!previous || String(row.updated_at || '').localeCompare(String(previous.updated_at || '')) >= 0) {
      latest[row.note_id] = row;
    }
  });
  return Object.values(latest)
    .filter(row => includeArchived || String(row.archived).toLowerCase() !== 'true')
    .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
}

function latestEditTypes_() {
  const sheet = sheet_(EDITS_SHEET, editHeaders_());
  const latest = {};
  rowsAsObjects_(sheet).forEach(row => {
    if (!row.note_id) return;
    let payload = {};
    try {
      payload = JSON.parse(row.payload || '{}');
    } catch (error) {
      payload = {};
    }
    const noteType = row.note_type || payload.note_type || '';
    if (!noteType) return;
    const previous = latest[row.note_id];
    if (!previous || String(row.updated_at || '').localeCompare(String(previous.updated_at || '')) >= 0) {
      latest[row.note_id] = { note_type: noteType, updated_at: row.updated_at || '' };
    }
  });
  return Object.keys(latest).reduce((acc, noteId) => {
    acc[noteId] = latest[noteId].note_type;
    return acc;
  }, {});
}

function getEdit_(noteId) {
  if (!noteId) return {};
  const sheet = sheet_(EDITS_SHEET, editHeaders_());
  const row = rowsAsObjects_(sheet)
    .filter(item => String(item.note_id) === String(noteId))
    .sort((a, b) => String(b.updated_at || '').localeCompare(String(a.updated_at || '')))[0];
  if (!row) return {};
  let payload = {};
  try {
    payload = JSON.parse(row.payload || '{}');
  } catch (error) {
    payload = {};
  }
  return { note_id: row.note_id, note_type: row.note_type || payload.note_type || '', payload, updated_at: row.updated_at };
}

function listStatus_() {
  const sheet = sheet_(STATUS_SHEET, statusHeaders_());
  return rowsAsObjects_(sheet);
}
