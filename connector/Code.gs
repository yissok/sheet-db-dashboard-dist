/**
 * sheet-db 0.2.0 — private current-state tables with deferred direct-edit reconciliation.
 * @OnlyCurrentDoc
 * Manual spreadsheet editing does not honor Apps Script locks. Checks detect observable
 * changes; they do not establish transactions with concurrent human editing.
 */
const SHEET_DB_PROOF_ = Object.freeze({
  protocol: 2, schema: 2, version: '0.2.0', sheetName: 'records',
  metaName: '_sheetdb_meta', opsName: '_sheetdb_ops',
  metaHeader: 'sheet-db storage schema 2', opsHeader: 'sheet-db operations schema 2',
  sessionKey: 'sheetdb.connection.v2', initializedKey: 'sheetdb.initialized.v2', initializingKey: 'sheetdb.initializing.v2',
  sessionDuration: 30 * 60 * 1000, receiptDuration: 24 * 60 * 60 * 1000,
  maxReceipts: 1000, maxRecords: 1000,
  headers: ['id', 'label', '_revision', '_status'], types: ['id', 'text', 'integer', 'status'],
  opHeaders: ['operationId', 'issuedAt', 'action', 'recordId', 'requestHash', 'priorRevision', 'priorHash', 'targetRevision', 'targetHash', 'status', 'deleted', 'row']
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('sheet-db').addItem('Open connection panel', 'openSheetDbPanel').addToUi();
}
function openSheetDbPanel() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('Panel').setTitle('sheet-db connection'));
}

/** Best-effort marker only. A full reconciliation runs before every adapter request. */
function onEdit(event) {
  try {
    if (!event || !event.range || event.range.getSheet().getName() !== 'records') return;
    const first = Math.max(3, event.range.getRow());
    const last = event.range.getRow() + event.range.getNumRows() - 1;
    if (last < first) return;
    event.range.getSheet().getRange(first, 4, last - first + 1, 1)
      .setValues(Array.from({ length: last - first + 1 }, function () { return ['Pending']; }));
  } catch (_) { /* Trigger queue/delivery is not an authoritative change log. */ }
}

function sheetDbPrepareConnection(input) {
  return sheetDbSafe_(function () {
    sheetDbAssert_(input && typeof input === 'object', 'VALIDATION_ERROR', 'Connection information is missing.');
    sheetDbAssert_(typeof input.channel === 'string' && /^[A-Za-z0-9_-]{43}$/.test(input.channel), 'VALIDATION_ERROR', 'The connection code is invalid.');
    sheetDbAssert_(typeof input.appOrigin === 'string' && sheetDbValidOrigin_(input.appOrigin), 'VALIDATION_ERROR', 'Use an HTTPS app address, or localhost for development.');
    const now = Date.now();
    sheetDbAssert_(Number.isFinite(input.expiresAt) && input.expiresAt > now && input.expiresAt <= now + 5 * 60 * 1000,
      'CONNECTION_EXPIRED', 'This connection invitation has expired. Open the app again.');
    const identity = sheetDbIdentity_();
    const sessionExpiresAt = now + SHEET_DB_PROOF_.sessionDuration;
    PropertiesService.getUserProperties().setProperty(SHEET_DB_PROOF_.sessionKey, JSON.stringify({
      channel: input.channel, appOrigin: input.appOrigin, accountId: identity.accountId,
      databaseId: identity.databaseId, expiresAt: sessionExpiresAt
    }));
    return { identity: identity, sessionExpiresAt: sessionExpiresAt };
  });
}

function sheetDbExecute(request) {
  return sheetDbSafe_(function () {
    sheetDbAssert_(request && typeof request === 'object', 'VALIDATION_ERROR', 'The request is missing.');
    sheetDbAssert_(request.protocol === 2, 'PROTOCOL_MISMATCH', 'Update the app and connector to compatible versions.');
    const identity = sheetDbIdentity_();
    const encoded = PropertiesService.getUserProperties().getProperty(SHEET_DB_PROOF_.sessionKey);
    let session;
    try { session = encoded ? JSON.parse(encoded) : null; } catch (_) { session = null; }
    sheetDbAssert_(session && session.channel === request.channel && session.expiresAt > Date.now(), 'CONNECTION_EXPIRED', 'Open the connection panel and connect the app again.');
    sheetDbAssert_(session.accountId === identity.accountId && session.databaseId === identity.databaseId && request.context &&
      request.context.accountId === identity.accountId && request.context.databaseId === identity.databaseId,
      'IDENTITY_MISMATCH', 'The account or database changed. Reconnect before trying again.');
    const command = sheetDbCommand_(request.command);
    const lock = LockService.getDocumentLock();
    sheetDbAssert_(lock && lock.tryLock(5000), 'BUSY', 'The database is busy. Retry the same request.', { retryable: true });
    try {
      const storage = sheetDbStorage_();
      sheetDbRecover_(storage);
      sheetDbPrune_(storage);
      const state = sheetDbReconcile_(storage);
      if (command.action === 'list') return { result: state.rows.map(sheetDbRecord_) };
      if (command.action === 'read') {
        const found = state.byId[command.recordId];
        sheetDbAssert_(found, 'NOT_FOUND', 'This record does not exist.');
        return { result: sheetDbRecord_(found) };
      }
      return { result: sheetDbMutate_(storage, state, command) };
    } finally { lock.releaseLock(); }
  });
}

function sheetDbIdentity_() {
  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  sheetDbAssert_(effective && /^[^\s@]+@[^\s@]+$/.test(effective), 'IDENTITY_UNAVAILABLE', 'Google did not identify the authorized account. Use a browser profile signed into only the Sheet owner.');
  sheetDbAssert_(!active || active === effective, 'IDENTITY_MISMATCH', 'Google returned different active and authorized accounts. Use a browser profile signed into only the Sheet owner.');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheetDbAssert_(spreadsheet, 'IDENTITY_UNAVAILABLE', 'Open this connector from its own Google Sheet.');
  return { accountId: effective, accountEmail: effective, databaseId: spreadsheet.getId(), databaseName: spreadsheet.getName(), protocolVersion: 2, schemaVersion: 2, connectorVersion: '0.2.0' };
}

function sheetDbCommand_(input) {
  sheetDbAssert_(input && typeof input === 'object' && !Array.isArray(input), 'VALIDATION_ERROR', 'The command is invalid.');
  const action = input.action;
  sheetDbAssert_(['create', 'read', 'update', 'delete', 'list'].indexOf(action) !== -1, 'VALIDATION_ERROR', 'This command is not supported.');
  const allowed = { action: true }, command = { action: action };
  if (['create', 'update', 'delete'].indexOf(action) !== -1) {
    allowed.operationId = allowed.issuedAt = true;
    sheetDbAssert_(typeof input.operationId === 'string' && /^[A-Za-z0-9_-]{16,128}$/.test(input.operationId), 'VALIDATION_ERROR', 'A stable operation code is required.');
    sheetDbAssert_(Number.isSafeInteger(input.issuedAt), 'VALIDATION_ERROR', 'The original operation time is required.');
    sheetDbAssert_(input.issuedAt > Date.now() - SHEET_DB_PROOF_.receiptDuration && input.issuedAt <= Date.now() + 60000,
      'RETRY_EXPIRED', 'This operation is too old or has an invalid time. Read the current database before creating a new operation.');
    command.operationId = input.operationId; command.issuedAt = input.issuedAt;
  }
  if (action !== 'list') {
    allowed.recordId = true;
    sheetDbAssert_(typeof input.recordId === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(input.recordId), 'VALIDATION_ERROR', 'The record identifier is invalid.');
    command.recordId = input.recordId;
  }
  if (action === 'update' || action === 'delete') {
    allowed.expectedRevision = true;
    sheetDbAssert_(Number.isSafeInteger(input.expectedRevision) && input.expectedRevision >= 1, 'VALIDATION_ERROR', 'The current record revision is required.');
    command.expectedRevision = input.expectedRevision;
  }
  if (action === 'create' || action === 'update') {
    allowed.data = true;
    sheetDbAssert_(input.data && typeof input.data === 'object' && !Array.isArray(input.data) && Object.keys(input.data).length === 1 &&
      typeof input.data.label === 'string' && input.data.label.trim() && input.data.label.length <= 500 && !/^=/.test(input.data.label),
      'VALIDATION_ERROR', 'Enter a text label between 1 and 500 characters. Formula expressions are not supported.');
    command.data = { label: input.data.label };
  }
  sheetDbAssert_(Object.keys(input).every(function (key) { return Object.prototype.hasOwnProperty.call(allowed, key); }), 'VALIDATION_ERROR', 'The command contains unsupported fields.');
  return command;
}

function sheetDbStorage_() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  sheetDbAssert_(!book.getSheetByName('_sheetdb_proof_v1'), 'UPGRADE_REQUIRED',
    'This Sheet contains the old audit journal. Keep it unchanged and install connector 0.2.0 in a new blank private Sheet. Automatic migration is not available.');
  let table = book.getSheetByName('records'), meta = book.getSheetByName('_sheetdb_meta'), ops = book.getSheetByName('_sheetdb_ops');
  const properties = PropertiesService.getDocumentProperties();
  const initialized = properties.getProperty(SHEET_DB_PROOF_.initializedKey);
  if (!table && !meta && !ops && initialized !== book.getId()) {
    // Persist ownership of empty initialization before the first worksheet mutation.
    properties.setProperty(SHEET_DB_PROOF_.initializingKey, book.getId());
    sheetDbAssert_(properties.getProperty(SHEET_DB_PROOF_.initializingKey) === book.getId(),
      'SCHEMA_CHANGED', 'Database setup could not be started safely. Retry in a new blank private Sheet.');
  }
  if (properties.getProperty(SHEET_DB_PROOF_.initializingKey) === book.getId()) {
    const initializedStorage = sheetDbInitialize_(book, properties);
    table = initializedStorage.table; meta = initializedStorage.meta; ops = initializedStorage.ops;
  }
  sheetDbAssert_(table && meta && ops, 'SCHEMA_CHANGED', 'A managed table or supporting tab was removed. Restore its original structure before continuing.');
  sheetDbHeaders_(table, [SHEET_DB_PROOF_.headers, SHEET_DB_PROOF_.types], 4, 'SCHEMA_CHANGED');
  sheetDbAssert_(meta.getRange(1, 1).getValue() === SHEET_DB_PROOF_.metaHeader && ops.getRange(1, 1).getValue() === SHEET_DB_PROOF_.opsHeader,
    'SCHEMA_CHANGED', 'The storage schema was changed. Restore the supporting tabs.');
  sheetDbHeaders_(meta, [['id', 'revision', 'hash']], 3, 'SCHEMA_CHANGED', 2);
  sheetDbHeaders_(ops, [SHEET_DB_PROOF_.opHeaders], 12, 'SCHEMA_CHANGED', 2);
  const storage = { table: table, meta: meta, ops: ops };
  if (meta.getRange(1, 2).getValue() !== book.getId()) sheetDbAdoptCopy_(storage, book.getId());
  properties.setProperty(SHEET_DB_PROOF_.initializedKey, book.getId());
  return storage;
}

/** Resume only our own empty, interrupted setup. Never adopt user records as setup data. */
function sheetDbInitialize_(book, properties) {
  const definitions = [
    { name: 'records', expected: [SHEET_DB_PROOF_.headers, SHEET_DB_PROOF_.types], columns: 4 },
    { name: '_sheetdb_meta', expected: [[SHEET_DB_PROOF_.metaHeader, book.getId(), ''], ['id', 'revision', 'hash']], columns: 3 },
    { name: '_sheetdb_ops', expected: [[SHEET_DB_PROOF_.opsHeader].concat(Array(11).fill('')), SHEET_DB_PROOF_.opHeaders], columns: 12 }
  ];
  // Check every existing tab before creating/filling anything. Unknown data stays untouched.
  definitions.forEach(function (definition) {
    const sheet = book.getSheetByName(definition.name);
    if (sheet) sheetDbInitializationEmpty_(sheet, definition);
  });
  const sheets = Object.create(null);
  definitions.forEach(function (definition) {
    const sheet = book.getSheetByName(definition.name) || book.insertSheet(definition.name);
    sheetDbInitializationEmpty_(sheet, definition);
    sheet.getRange(1, 1, 2, definition.columns).setValues(definition.expected);
    sheets[definition.name] = sheet;
  });
  const table = sheets.records, meta = sheets._sheetdb_meta, ops = sheets._sheetdb_ops;
  if (table.getMaxRows() < SHEET_DB_PROOF_.maxRecords + 2)
    table.insertRowsAfter(table.getMaxRows(), SHEET_DB_PROOF_.maxRecords + 2 - table.getMaxRows());
  table.getRange(3, 2, table.getMaxRows() - 2, 1).setNumberFormat('@');
  if (table.setFrozenRows) table.setFrozenRows(2);
  sheetDbProtect_(table.getRange(1, 1, 2, 4), 'Field names and types are managed by sheet-db.');
  [1, 3, 4].forEach(function (column) {
    sheetDbProtect_(table.getRange(3, column, table.getMaxRows() - 2, 1), 'Managed sheet-db ID, revision or status. Edit record values instead.');
  });
  if (meta.hideSheet) meta.hideSheet(); if (ops.hideSheet) ops.hideSheet();
  SpreadsheetApp.flush();
  definitions.forEach(function (definition) {
    const sheet = sheets[definition.name];
    sheetDbHeaders_(sheet, definition.expected, definition.columns, 'SCHEMA_CHANGED');
    sheetDbAssert_(sheet.getLastRow() === 2, 'SCHEMA_CHANGED',
      'A record was added before database setup finished. Keep this Sheet unchanged and use a new blank private Sheet.');
  });
  properties.setProperty(SHEET_DB_PROOF_.initializedKey, book.getId());
  properties.deleteProperty(SHEET_DB_PROOF_.initializingKey);
  return { table: table, meta: meta, ops: ops };
}
function sheetDbInitializationEmpty_(sheet, definition) {
  const range = sheet.getRange(1, 1, 2, definition.columns);
  const values = range.getValues();
  sheetDbAssert_(sheet.getLastRow() <= 2 && sheet.getLastColumn() <= definition.columns &&
    !range.getFormulas().some(function (line) { return line.some(Boolean); }) &&
    values.every(function (line, row) { return line.every(function (value, column) {
      return value === '' || value === definition.expected[row][column];
    }); }), 'SCHEMA_CHANGED',
    'Setup was interrupted and the Sheet now contains unexpected data. Nothing has been overwritten; preserve this Sheet and use a new blank private Sheet.');
}

function sheetDbProtect_(range, description) {
  if (!range.protect) return;
  const protection = range.protect();
  if (protection.setDescription) protection.setDescription(description);
  protection.setWarningOnly(true);
}

function sheetDbAdoptCopy_(storage, databaseId) {
  const receipts = sheetDbReceipts_(storage);
  sheetDbAssert_(!receipts.some(function (receipt) { return receipt.status === 'PREPARED'; }), 'COPY_REQUIRES_REVIEW',
    'This copy contains an interrupted operation from another database. Preserve it for inspection; no copied operation has been replayed. Use a fresh blank Sheet for a new database.');
  const rows = sheetDbReadRows_(storage), baseline = sheetDbBaselines_(storage), seen = Object.create(null);
  sheetDbAssert_(rows.length === Object.keys(baseline).length && rows.every(function (row) {
    const accepted = baseline[row.id];
    if (!accepted || seen[row.id] || row.formula || row.revision !== accepted.revision || row.hash !== accepted.hash) return false;
    seen[row.id] = true; return true;
  }), 'COPY_REQUIRES_REVIEW', 'This copy contains pending spreadsheet edits. Accept them in the original database before making a new copy, or use a fresh blank Sheet.');
  sheetDbStable_(storage, rows);
  receipts.reverse().forEach(function (receipt) { storage.ops.deleteRow(receipt.storageRow); });
  SpreadsheetApp.flush();
  sheetDbStable_(storage, rows);
  storage.meta.getRange(1, 2).setValue(databaseId); SpreadsheetApp.flush();
  sheetDbAssert_(storage.meta.getRange(1, 2).getValue() === databaseId, 'COPY_REQUIRES_REVIEW', 'The copied database identity could not be verified.');
}

function sheetDbHeaders_(sheet, expected, columns, code, row) {
  const range = sheet.getRange(row || 1, 1, expected.length, columns);
  const values = range.getValues();
  sheetDbAssert_(JSON.stringify(values) === JSON.stringify(expected) && !range.getFormulas().some(function (line) { return line.some(Boolean); }) && sheet.getLastColumn() <= columns,
    code, 'Table names, field names and field types must stay unchanged in this version.');
}
function sheetDbHash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(function (byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}
function sheetDbCells_(sheet, columns) {
  const count = Math.max(0, sheet.getLastRow() - 2);
  return count ? sheet.getRange(3, 1, count, columns).getValues() : [];
}
function sheetDbWrite_(sheet, row, values) {
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
  if (sheet.getName() === 'records') sheet.getRange(row, 2).setNumberFormat('@');
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  SpreadsheetApp.flush();
  sheetDbAssert_(JSON.stringify(sheet.getRange(row, 1, 1, values.length).getValues()[0]) === JSON.stringify(values),
    'WRITE_UNCONFIRMED', 'The write could not be verified. Retry the original operation after reconnecting.');
}
function sheetDbBaselines_(storage) {
  const result = Object.create(null);
  sheetDbCells_(storage.meta, 3).forEach(function (values, index) {
    if (values.every(function (value) { return value === ''; })) return;
    sheetDbAssert_(typeof values[0] === 'string' && values[0] && !result[values[0]] && Number.isSafeInteger(values[1]) && values[1] >= 1 && /^[a-f0-9]{64}$/.test(values[2]),
      'STORAGE_CORRUPT', 'The current record metadata is invalid.');
    result[values[0]] = { id: values[0], revision: values[1], hash: values[2], row: index + 3 };
  });
  return result;
}
function sheetDbSetBaseline_(storage, id, revision, hash) {
  const current = sheetDbBaselines_(storage)[id];
  sheetDbWrite_(storage.meta, current ? current.row : Math.max(3, storage.meta.getLastRow() + 1), [id, revision, hash]);
}
function sheetDbRemoveBaseline_(storage, id) {
  const current = sheetDbBaselines_(storage)[id];
  if (current) { storage.meta.deleteRow(current.row); SpreadsheetApp.flush(); }
}
function sheetDbReadRows_(storage) {
  const values = sheetDbCells_(storage.table, 4), result = [];
  const formulas = values.length ? storage.table.getRange(3, 1, values.length, 4).getFormulas() : [];
  values.forEach(function (cells, index) {
    if (cells.slice(0, 3).every(function (value) { return value === ''; })) return;
    result.push({ row: index + 3, id: cells[0], label: cells[1], revision: cells[2], status: cells[3],
      hash: typeof cells[1] === 'string' ? sheetDbHash_(cells[1]) : '', formula: (formulas[index] || []).some(Boolean) });
  });
  return result;
}
function sheetDbRecord_(row) { return { id: row.id, revision: row.revision, data: { label: row.label }, deleted: false }; }
function sheetDbFingerprint_(rows) {
  return sheetDbHash_(JSON.stringify(rows.map(function (row) { return [row.row, row.id, row.label, row.revision, row.formula]; })));
}
function sheetDbStable_(storage, expected) {
  sheetDbAssert_(sheetDbFingerprint_(sheetDbReadRows_(storage)) === sheetDbFingerprint_(expected), 'MANUAL_EDIT_IN_PROGRESS',
    'The spreadsheet changed while this request was being checked. Finish editing and retry the same request.');
}
function sheetDbIssue_(storage, row, field, message, issues) {
  issues.push({ row: row.row, field: field, message: message });
  try {
    storage.table.getRange(row.row, 4).setValue('Needs attention');
    const range = storage.table.getRange(row.row, field === 'label' ? 2 : field === '_revision' ? 3 : 1);
    if (range.setBackground) range.setBackground('#fff0d6');
    if (range.setNote) range.setNote(message);
  } catch (_) { /* The structured validation error remains authoritative. */ }
}

function sheetDbReconcile_(storage) {
  sheetDbCells_(storage.table, 4).forEach(function (cells, index) {
    if (cells.slice(0, 3).every(function (value) { return value === ''; }) && cells[3])
      storage.table.getRange(index + 3, 4).setValue('');
  });
  let rows = sheetDbReadRows_(storage);
  const baseline = sheetDbBaselines_(storage), byId = Object.create(null), issues = [];
  const missing = Object.keys(baseline).filter(function (id) { return !rows.some(function (row) { return row.id === id; }); });
  rows.forEach(function (row) {
    if (row.formula || typeof row.label !== 'string' || !row.label.trim() || row.label.length > 500 || /^=/.test(row.label))
      sheetDbIssue_(storage, row, 'label', 'Use text between 1 and 500 characters; formulas are not supported.', issues);
    if (row.id === '') {
      if (row.revision !== '') sheetDbIssue_(storage, row, '_revision', 'Leave managed revisions blank on a new row.', issues);
    } else {
      if (typeof row.id !== 'string' || !baseline[row.id]) sheetDbIssue_(storage, row, 'id', 'Do not edit existing IDs. Leave the ID blank when adding a new row.', issues);
      if (byId[row.id]) sheetDbIssue_(storage, row, 'id', 'Two rows have the same ID. Restore each original ID.', issues);
      if (baseline[row.id] && row.revision !== baseline[row.id].revision)
        sheetDbIssue_(storage, row, '_revision', 'Do not edit the managed revision.', issues);
      byId[row.id] = row;
    }
  });
  sheetDbAssert_(rows.length <= SHEET_DB_PROOF_.maxRecords, 'LIMIT_EXCEEDED', 'This release supports at most 1,000 current records.');
  sheetDbAssert_(!issues.length, 'MANUAL_EDIT_INVALID', 'Fix the highlighted spreadsheet cells before this request can continue.', { issues: issues });
  sheetDbStable_(storage, rows);
  // A missing physical row is a deletion; its current metadata is removed, not archived.
  missing.forEach(function (id) { sheetDbRemoveBaseline_(storage, id); });
  rows.forEach(function (row) {
    const previous = row.id ? baseline[row.id] : null;
    if (!previous || previous.hash !== row.hash) {
      sheetDbStable_(storage, rows);
      const id = previous ? row.id : Utilities.getUuid();
      const receipt = {
        operationId: Utilities.getUuid(), issuedAt: Date.now(), action: previous ? 'manual-update' : 'manual-create', recordId: id,
        requestHash: row.hash, priorRevision: previous ? previous.revision : 0, priorHash: previous ? previous.hash : '',
        targetRevision: previous ? previous.revision + 1 : 1, targetHash: row.hash, status: 'PREPARED', deleted: false, row: row.row
      };
      sheetDbAssert_(sheetDbReceipts_(storage).length < SHEET_DB_PROOF_.maxReceipts, 'LIMIT_EXCEEDED',
        'The 24-hour retry receipt limit has been reached. Pending spreadsheet edits are unchanged; wait for old receipts to expire.');
      try {
        sheetDbSaveReceipt_(storage, receipt);
        // Only managed cells change; the user's actual label is never overwritten here.
        storage.table.getRange(row.row, 1).setValue(id);
        storage.table.getRange(row.row, 3, 1, 2).setValues([[receipt.targetRevision, '']]);
        SpreadsheetApp.flush();
        const current = sheetDbReadRows_(storage).filter(function (entry) { return entry.id === id; });
        sheetDbAssert_(current.length === 1 && current[0].hash === receipt.targetHash && current[0].revision === receipt.targetRevision,
          'RECOVERY_REQUIRED', 'A direct edit changed while it was being accepted. Keep the Sheet unchanged for inspection.');
        sheetDbSetBaseline_(storage, id, receipt.targetRevision, receipt.targetHash);
        sheetDbDeleteReceipt_(storage, receipt);
      } catch (error) {
        if (error.sheetDbCode === 'RECOVERY_REQUIRED') throw error;
        sheetDbFail_('RECOVERY_REQUIRED', 'Accepting a direct edit was interrupted. Keep the Sheet unchanged and retry to check recovery.');
      }
      rows = sheetDbReadRows_(storage);
    }
    sheetDbClearMarks_(storage, row.row);
  });
  rows = sheetDbReadRows_(storage);
  sheetDbStable_(storage, rows);
  const finalBaseline = sheetDbBaselines_(storage), finalMap = Object.create(null);
  rows.forEach(function (row) {
    const accepted = finalBaseline[row.id];
    sheetDbAssert_(accepted && row.revision === accepted.revision && row.hash === accepted.hash,
      'MANUAL_EDIT_IN_PROGRESS', 'The spreadsheet changed during reconciliation. Finish editing and retry.');
    finalMap[row.id] = row;
  });
  SpreadsheetApp.flush();
  return { rows: rows, byId: finalMap };
}

function sheetDbClearMarks_(storage, row) {
  storage.table.getRange(row, 4).setValue('');
  const cells = storage.table.getRange(row, 1, 1, 3);
  if (cells.setBackground) cells.setBackground(null);
  if (cells.setNote) cells.setNote('');
}

function sheetDbReceiptValues_(receipt) {
  return SHEET_DB_PROOF_.opHeaders.map(function (key) { return receipt[key] === undefined ? '' : receipt[key]; });
}
function sheetDbReceipts_(storage) {
  return sheetDbCells_(storage.ops, 12).map(function (values, index) {
    const receipt = { storageRow: index + 3 };
    SHEET_DB_PROOF_.opHeaders.forEach(function (key, column) { receipt[key] = values[column]; });
    sheetDbAssert_(receipt.operationId && ['PREPARED', 'APPLIED'].indexOf(receipt.status) !== -1 && Number.isSafeInteger(receipt.issuedAt),
      'STORAGE_CORRUPT', 'The retry metadata is invalid.');
    return receipt;
  });
}
function sheetDbSaveReceipt_(storage, receipt) {
  const receipts = sheetDbReceipts_(storage);
  const existing = receipts.filter(function (entry) { return entry.operationId === receipt.operationId; })[0];
  if (existing) receipt.storageRow = existing.storageRow;
  else {
    sheetDbAssert_(receipts.length < SHEET_DB_PROOF_.maxReceipts, 'LIMIT_EXCEEDED',
      'The 24-hour retry receipt limit has been reached. Wait for old receipts to expire before making another change.');
    receipt.storageRow = Math.max(3, storage.ops.getLastRow() + 1);
  }
  sheetDbWrite_(storage.ops, receipt.storageRow, sheetDbReceiptValues_(receipt));
}
function sheetDbDeleteReceipt_(storage, receipt) {
  const current = sheetDbReceipts_(storage).filter(function (entry) { return entry.operationId === receipt.operationId; })[0];
  if (current) { storage.ops.deleteRow(current.storageRow); SpreadsheetApp.flush(); }
}
function sheetDbAck_(receipt) {
  return { id: receipt.recordId, revision: receipt.targetRevision, operationId: receipt.operationId, applied: true, deleted: receipt.deleted === true };
}
function sheetDbPrune_(storage) {
  sheetDbReceipts_(storage).reverse().forEach(function (receipt) {
    if (receipt.status === 'APPLIED' && receipt.issuedAt <= Date.now() - SHEET_DB_PROOF_.receiptDuration)
      storage.ops.deleteRow(receipt.storageRow);
  });
  SpreadsheetApp.flush();
}

function sheetDbRecover_(storage) {
  sheetDbReceipts_(storage).forEach(function (receipt) {
    if (receipt.status !== 'PREPARED') return;
    const rows = sheetDbReadRows_(storage);
    let current = rows.filter(function (row) { return row.id === receipt.recordId; });
    const manual = /^manual-/.test(receipt.action);
    if (receipt.action === 'manual-create' && !current.length) {
      // The ID assignment may not have run. Only accept the original unchanged physical row.
      const candidate = rows.filter(function (row) { return row.row === receipt.row && row.id === '' && row.revision === '' && row.hash === receipt.targetHash && !row.formula; });
      sheetDbAssert_(candidate.length === 1, 'RECOVERY_REQUIRED', 'An interrupted new spreadsheet row is ambiguous. Keep the Sheet unchanged for inspection.');
      storage.table.getRange(candidate[0].row, 1).setValue(receipt.recordId);
      storage.table.getRange(candidate[0].row, 3).setValue(receipt.targetRevision);
      SpreadsheetApp.flush();
      current = sheetDbReadRows_(storage).filter(function (row) { return row.id === receipt.recordId; });
    }
    if (receipt.deleted === true) {
      sheetDbAssert_(!current.length, 'RECOVERY_REQUIRED', 'An interrupted deletion has an uncertain outcome. Keep the Sheet unchanged for inspection.');
      sheetDbRemoveBaseline_(storage, receipt.recordId);
    } else {
      sheetDbAssert_(current.length === 1 && !current[0].formula && current[0].hash === receipt.targetHash,
        'RECOVERY_REQUIRED', 'An interrupted write has an uncertain outcome. Keep the Sheet unchanged for inspection.');
      if (manual && (current[0].revision === receipt.priorRevision || (receipt.action === 'manual-create' && current[0].revision === ''))) {
        storage.table.getRange(current[0].row, 3).setValue(receipt.targetRevision); SpreadsheetApp.flush();
        current = sheetDbReadRows_(storage).filter(function (row) { return row.id === receipt.recordId; });
      }
      sheetDbAssert_(current[0].revision === receipt.targetRevision, 'RECOVERY_REQUIRED', 'The interrupted write revision is ambiguous.');
      sheetDbSetBaseline_(storage, receipt.recordId, receipt.targetRevision, receipt.targetHash);
    }
    if (manual) sheetDbDeleteReceipt_(storage, receipt);
    else { receipt.status = 'APPLIED'; sheetDbSaveReceipt_(storage, receipt); }
  });
}

function sheetDbMutate_(storage, state, command) {
  const signature = sheetDbHash_(JSON.stringify(command));
  const previous = sheetDbReceipts_(storage).filter(function (receipt) { return receipt.operationId === command.operationId; })[0];
  if (previous) {
    sheetDbAssert_(previous.requestHash === signature, 'IDEMPOTENCY_CONFLICT', 'This operation code was already used for different data.');
    sheetDbAssert_(previous.status === 'APPLIED', 'RECOVERY_REQUIRED', 'The earlier operation has an uncertain outcome.');
    return sheetDbAck_(previous);
  }
  const current = state.byId[command.recordId];
  if (command.action === 'create') {
    sheetDbAssert_(!current, 'ALREADY_EXISTS', 'This record ID already exists.');
    sheetDbAssert_(state.rows.length < SHEET_DB_PROOF_.maxRecords, 'LIMIT_EXCEEDED', 'This release supports at most 1,000 current records.');
  } else {
    sheetDbAssert_(current, 'NOT_FOUND', 'This record does not exist.');
    sheetDbAssert_(current.revision === command.expectedRevision, 'REVISION_CONFLICT', 'The record changed. Read its current revision before editing.', { currentRevision: current.revision });
  }
  sheetDbStable_(storage, state.rows);
  const receipt = {
    operationId: command.operationId, issuedAt: command.issuedAt, action: command.action, recordId: command.recordId,
    requestHash: signature, priorRevision: current ? current.revision : 0, priorHash: current ? current.hash : '',
    targetRevision: current ? current.revision + 1 : 1, targetHash: command.action === 'delete' ? '' : sheetDbHash_(command.data.label),
    status: 'PREPARED', deleted: command.action === 'delete', row: current ? current.row : Math.max(3, storage.table.getLastRow() + 1)
  };
  // Capacity is known before dispatch; a full receipt table is not an uncertain write.
  sheetDbAssert_(sheetDbReceipts_(storage).length < SHEET_DB_PROOF_.maxReceipts, 'LIMIT_EXCEEDED',
    'The 24-hour retry receipt limit has been reached. Wait for old receipts to expire.');
  // Persist a payload-free intent before touching the record. Ambiguous interruption fails closed.
  let attemptedRecordWrite = false;
  try {
    sheetDbSaveReceipt_(storage, receipt);
    sheetDbStable_(storage, state.rows);
    attemptedRecordWrite = true;
    if (receipt.deleted) { storage.table.deleteRow(current.row); SpreadsheetApp.flush(); }
    else sheetDbWrite_(storage.table, receipt.row, [receipt.recordId, command.data.label, receipt.targetRevision, '']);
    const confirmed = sheetDbReadRows_(storage).filter(function (row) { return row.id === receipt.recordId; });
    sheetDbAssert_(receipt.deleted ? !confirmed.length : confirmed.length === 1 && !confirmed[0].formula &&
      confirmed[0].hash === receipt.targetHash && confirmed[0].revision === receipt.targetRevision,
      'WRITE_UNCONFIRMED', 'The record write could not be verified.');
    if (receipt.deleted) sheetDbRemoveBaseline_(storage, receipt.recordId);
    else sheetDbSetBaseline_(storage, receipt.recordId, receipt.targetRevision, receipt.targetHash);
    receipt.status = 'APPLIED'; sheetDbSaveReceipt_(storage, receipt);
    return sheetDbAck_(receipt);
  } catch (error) {
    if (!attemptedRecordWrite && error.sheetDbCode === 'MANUAL_EDIT_IN_PROGRESS') {
      // We observed this change before attempting any row mutation, so cancel the intent.
      sheetDbDeleteReceipt_(storage, receipt);
      throw error;
    }
    sheetDbFail_('WRITE_UNCONFIRMED', 'The write may have completed. Retry the same operation to check it; do not create a new operation ID.',
      { operationId: command.operationId, recordId: command.recordId, retryable: true, cause: error.sheetDbCode || 'GOOGLE_ERROR' });
  }
}

function sheetDbValidOrigin_(origin) {
  return /^(?:https:\/\/[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?(?::[0-9]{1,5})?|http:\/\/(?:localhost|127\.0\.0\.1|\[::1\])(?::[0-9]{1,5})?)$/.test(origin);
}
function sheetDbSafe_(work) {
  try { return work(); } catch (error) {
    if (error && error.sheetDbCode) {
      const structured = { code: error.sheetDbCode, message: error.message };
      if (error.sheetDbDetails) structured.details = error.sheetDbDetails;
      return { error: structured };
    }
    const message = String(error && error.message || 'Google could not complete this request.').slice(0, 400);
    return { error: { code: /quota|too many times|limit exceeded/i.test(message) ? 'QUOTA_EXCEEDED' : 'GOOGLE_ERROR', message: message, details: { retryable: false } } };
  }
}
function sheetDbAssert_(condition, code, message, details) { if (!condition) sheetDbFail_(code, message, details); }
function sheetDbFail_(code, message, details) {
  const error = new Error(message); error.sheetDbCode = code; error.sheetDbDetails = details; throw error;
}
