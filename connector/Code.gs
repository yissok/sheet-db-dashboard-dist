/**
 * sheet-db 0.3.0 — private current-state tables with deferred direct-edit reconciliation.
 * @OnlyCurrentDoc
 * Manual spreadsheet editing does not honor Apps Script locks. Checks detect observable
 * changes; they do not establish transactions with concurrent human editing.
 */
const SHEET_DB_PROOF_ = Object.freeze({
  protocol: 2, schema: 2, version: '0.3.0', sheetName: 'records',
  metaName: '_sheetdb_meta', opsName: '_sheetdb_ops',
  metaHeader: 'sheet-db storage schema 2', opsHeader: 'sheet-db operations schema 2',
  sessionKey: 'sheetdb.connection.v2', initializedKey: 'sheetdb.initialized.v2', initializingKey: 'sheetdb.initializing.v2',
  sessionDuration: 30 * 60 * 1000, receiptDuration: 24 * 60 * 60 * 1000,
  maxReceipts: 1000, maxRecords: 1000,
  headers: ['id', 'label', '_revision', '_status'], types: ['id', 'text', 'integer', 'status'],
  opHeaders: ['operationId', 'issuedAt', 'action', 'recordId', 'requestHash', 'priorRevision', 'priorHash', 'targetRevision', 'targetHash', 'status', 'deleted', 'row']
});

let SHEET_DB_CURRENT_ = null;
let SHEET_DB_STATES_ = null;
// A spreadsheet-submit trigger supplies its own workbook; the panel uses the active one.
let SHEET_DB_BOOK_ = null;
function sheetDbBook_() { return SHEET_DB_BOOK_ || SpreadsheetApp.getActiveSpreadsheet(); }
function sheetDbConfig_() { return SHEET_DB_CURRENT_ || SHEET_DB_PROOF_; }

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('sheet-db');
  if (typeof sheetDbOpenSetup === 'function') menu.addItem('Open setup', 'sheetDbOpenSetup');
  if (typeof sheetDbSetupFormsPoc === 'function') {
    menu.addItem('Set up public Forms POC', 'sheetDbSetupFormsPoc')
      .addItem('Show public connection details', 'sheetDbShowFormsConnection');
  }
  menu.addItem('Open connection panel', 'openSheetDbPanel').addToUi();
}
function openSheetDbPanel() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('Panel').setTitle('sheet-db connection'));
}

/** Best-effort marker only. A full reconciliation runs before every adapter request. */
function onEdit(event) {
  if (typeof sheetDbSetupOnEdit_ === 'function' && sheetDbSetupOnEdit_(event)) return;
  try {
    if (!event || !event.range) return;
    const name = event.range.getSheet().getName();
    if (name !== 'records') {
      const schema = sheetDbLoadSchema_();
      const table = schema && schema.tables.find(function (item) { return item.name === name; });
      if (!table) return;
      SHEET_DB_CURRENT_ = sheetDbTableConfig_(table);
    }
    const first = Math.max(3, event.range.getRow());
    const last = event.range.getRow() + event.range.getNumRows() - 1;
    if (last >= first) event.range.getSheet().getRange(first, sheetDbConfig_().headers.length, last - first + 1, 1)
      .setValues(Array.from({ length: last - first + 1 }, function () { return ['Pending']; }));
  } catch (_) { /* Markers are best effort; each operation scans the tables. */ }
  finally { SHEET_DB_CURRENT_ = null; }
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
    const sessionExpiresAt = now + sheetDbConfig_().sessionDuration;
    PropertiesService.getUserProperties().setProperty(sheetDbConfig_().sessionKey, JSON.stringify({
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
    const encoded = PropertiesService.getUserProperties().getProperty(sheetDbConfig_().sessionKey);
    let session;
    try { session = encoded ? JSON.parse(encoded) : null; } catch (_) { session = null; }
    sheetDbAssert_(session && session.channel === request.channel && session.expiresAt > Date.now(), 'CONNECTION_EXPIRED', 'Open the connection panel and connect the app again.');
    sheetDbAssert_(session.accountId === identity.accountId && session.databaseId === identity.databaseId && request.context &&
      request.context.accountId === identity.accountId && request.context.databaseId === identity.databaseId,
      'IDENTITY_MISMATCH', 'The account or database changed. Reconnect before trying again.');
    return sheetDbExecuteCommand_(request.command);
  });
}

/** Trusted transports validate their context before entering the shared table engine. */
function sheetDbExecuteCommand_(input) {
  if (input && (input.action === 'ensureSchema' || input.table)) return sheetDbTableRequest_(input);
  const command = sheetDbCommand_(input);
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
}

function sheetDbIdentity_() {
  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  sheetDbAssert_(effective && /^[^\s@]+@[^\s@]+$/.test(effective), 'IDENTITY_UNAVAILABLE', 'Google did not identify the authorized account. Use a browser profile signed into only the Sheet owner.');
  sheetDbAssert_(!active || active === effective, 'IDENTITY_MISMATCH', 'Google returned different active and authorized accounts. Use a browser profile signed into only the Sheet owner.');
  const spreadsheet = sheetDbBook_();
  sheetDbAssert_(spreadsheet, 'IDENTITY_UNAVAILABLE', 'Open this connector from its own Google Sheet.');
  return { accountId: effective, accountEmail: effective, databaseId: spreadsheet.getId(), databaseName: spreadsheet.getName(), protocolVersion: 2, schemaVersion: 2, connectorVersion: '0.3.0', tableApiVersion: 1 };
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
    sheetDbAssert_(input.issuedAt > Date.now() - sheetDbConfig_().receiptDuration && input.issuedAt <= Date.now() + 60000,
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
    command.data = sheetDbValidateData_(input.data);

  }
  sheetDbAssert_(Object.keys(input).every(function (key) { return Object.prototype.hasOwnProperty.call(allowed, key); }), 'VALIDATION_ERROR', 'The command contains unsupported fields.');
  return command;
}

function sheetDbStorage_() {
  const book = sheetDbBook_();
  sheetDbAssert_(!book.getSheetByName('_sheetdb_proof_v1'), 'UPGRADE_REQUIRED',
    'This Sheet contains the old audit journal. Keep it unchanged and install connector 0.2.0 in a new blank private Sheet. Automatic migration is not available.');
  let table = book.getSheetByName(sheetDbConfig_().sheetName), meta = book.getSheetByName(sheetDbConfig_().metaName), ops = book.getSheetByName(sheetDbConfig_().opsName);
  const properties = PropertiesService.getDocumentProperties();
  const initialized = properties.getProperty(sheetDbConfig_().initializedKey);
  if (!table && !meta && !ops && initialized !== book.getId()) {
    // Persist ownership of empty initialization before the first worksheet mutation.
    properties.setProperty(sheetDbConfig_().initializingKey, book.getId());
    sheetDbAssert_(properties.getProperty(sheetDbConfig_().initializingKey) === book.getId(),
      'SCHEMA_CHANGED', 'Database setup could not be started safely. Retry in a new blank private Sheet.');
  }
  if (properties.getProperty(sheetDbConfig_().initializingKey) === book.getId()) {
    const initializedStorage = sheetDbInitialize_(book, properties);
    table = initializedStorage.table; meta = initializedStorage.meta; ops = initializedStorage.ops;
  }
  sheetDbAssert_(table && meta && ops, 'SCHEMA_CHANGED', 'A managed table or supporting tab was removed. Restore its original structure before continuing.');
  sheetDbHeaders_(table, [sheetDbConfig_().headers, sheetDbConfig_().types], sheetDbConfig_().headers.length, 'SCHEMA_CHANGED');
  sheetDbAssert_(meta.getRange(1, 1).getValue() === sheetDbConfig_().metaHeader && ops.getRange(1, 1).getValue() === sheetDbConfig_().opsHeader,
    'SCHEMA_CHANGED', 'The storage schema was changed. Restore the supporting tabs.');
  sheetDbHeaders_(meta, [['id', 'revision', 'hash']], 3, 'SCHEMA_CHANGED', 2);
  sheetDbHeaders_(ops, [sheetDbConfig_().opHeaders], 12, 'SCHEMA_CHANGED', 2);
  const storage = { table: table, meta: meta, ops: ops };
  if (meta.getRange(1, 2).getValue() !== book.getId()) sheetDbAdoptCopy_(storage, book.getId());
  properties.setProperty(sheetDbConfig_().initializedKey, book.getId());
  return storage;
}

/** Resume only our own empty, interrupted setup. Never adopt user records as setup data. */
function sheetDbInitialize_(book, properties) {
  const definitions = [
    { name: sheetDbConfig_().sheetName, expected: [sheetDbConfig_().headers, sheetDbConfig_().types], columns: sheetDbConfig_().headers.length },
    { name: sheetDbConfig_().metaName, expected: [[sheetDbConfig_().metaHeader, book.getId(), ''], ['id', 'revision', 'hash']], columns: 3 },
    { name: sheetDbConfig_().opsName, expected: [[sheetDbConfig_().opsHeader].concat(Array(11).fill('')), sheetDbConfig_().opHeaders], columns: 12 }
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
  const table = sheets[sheetDbConfig_().sheetName], meta = sheets[sheetDbConfig_().metaName], ops = sheets[sheetDbConfig_().opsName];
  if (table.getMaxRows() < sheetDbConfig_().maxRecords + 2)
    table.insertRowsAfter(table.getMaxRows(), sheetDbConfig_().maxRecords + 2 - table.getMaxRows());
  sheetDbFormatFields_(table, 3, table.getMaxRows() - 2);
  if (table.setFrozenRows) table.setFrozenRows(2);
  sheetDbProtect_(table.getRange(1, 1, 2, sheetDbConfig_().headers.length), 'Field names and types are managed by sheet-db.');
  [1, sheetDbConfig_().headers.length - 1, sheetDbConfig_().headers.length].forEach(function (column) {
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
  properties.setProperty(sheetDbConfig_().initializedKey, book.getId());
  properties.deleteProperty(sheetDbConfig_().initializingKey);
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
  if (sheet.getName() === sheetDbConfig_().sheetName) sheetDbFormatFields_(sheet, row, 1);
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
  const columns = sheetDbConfig_().headers.length;
  const values = sheetDbCells_(storage.table, columns), result = [];
  const formulas = values.length ? storage.table.getRange(3, 1, values.length, columns).getFormulas() : [];
  values.forEach(function (cells, index) {
    if (cells.slice(0, -1).every(function (value) { return value === ''; })) return;
    const data = {};
    sheetDbFields_().forEach(function (field, column) { data[field.name] = cells[column + 1]; });
    const label = SHEET_DB_CURRENT_ ? data : data.label;
    result.push({ row: index + 3, id: cells[0], label: label, data: data, revision: cells[columns - 2], status: cells[columns - 1],
      hash: sheetDbDataHash_(label), formula: (formulas[index] || []).some(Boolean) });
  });
  return result;
}
function sheetDbRecord_(row) { return { id: row.id, revision: row.revision, data: row.data || { label: row.label }, deleted: false }; }
function sheetDbFields_() { return sheetDbConfig_().fields || [{ name: 'label', type: 'text', required: true }]; }
function sheetDbDataHash_(value) { return sheetDbHash_(typeof value === 'string' ? value : JSON.stringify(value)); }
function sheetDbRowValues_(id, data, revision) {
  return [id].concat(sheetDbFields_().map(function (field) { return data[field.name]; }), [revision, '']);
}
function sheetDbFormatFields_(table, row, count) {
  sheetDbFields_().forEach(function (field, index) {
    if (['text', 'reference', 'date'].indexOf(field.type) !== -1) table.getRange(row, index + 2, count, 1).setNumberFormat('@');
  });
}

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
    storage.table.getRange(row.row, sheetDbConfig_().headers.length).setValue('Needs attention');
    const range = storage.table.getRange(row.row, field === '_revision' ? sheetDbConfig_().headers.length - 1 : Math.max(1, sheetDbConfig_().headers.indexOf(field) + 1));
    if (range.setBackground) range.setBackground('#fff0d6');
    if (range.setNote) range.setNote(message);
  } catch (_) { /* The structured validation error remains authoritative. */ }
}

function sheetDbReconcile_(storage) {
  sheetDbCells_(storage.table, sheetDbConfig_().headers.length).forEach(function (cells, index) {
    if (cells.slice(0, -1).every(function (value) { return value === ''; }) && cells[cells.length - 1])
      storage.table.getRange(index + 3, sheetDbConfig_().headers.length).setValue('');
  });
  let rows = sheetDbReadRows_(storage);
  const baseline = sheetDbBaselines_(storage), byId = Object.create(null), issues = [];
  const missing = Object.keys(baseline).filter(function (id) { return !rows.some(function (row) { return row.id === id; }); });
  rows.forEach(function (row) {
    try {
      sheetDbAssert_(!row.formula, 'VALIDATION_ERROR', 'Formulas are not supported in records.');
      sheetDbValidateData_(row.data);
    } catch (error) { sheetDbIssue_(storage, row, error.sheetDbDetails && error.sheetDbDetails.field || 'label', error.message, issues); }
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
  sheetDbAssert_(rows.length <= sheetDbConfig_().maxRecords, 'LIMIT_EXCEEDED', 'This table supports at most ' + sheetDbConfig_().maxRecords + ' current records.');
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
      sheetDbAssert_(sheetDbReceipts_(storage).length < sheetDbConfig_().maxReceipts, 'LIMIT_EXCEEDED',
        'The 24-hour retry receipt limit has been reached. Pending spreadsheet edits are unchanged; wait for old receipts to expire.');
      try {
        sheetDbSaveReceipt_(storage, receipt);
        // Only managed cells change; the user's actual label is never overwritten here.
        storage.table.getRange(row.row, 1).setValue(id);
        storage.table.getRange(row.row, sheetDbConfig_().headers.length - 1, 1, 2).setValues([[receipt.targetRevision, '']]);
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
    if (row.status || !previous || previous.hash !== row.hash) sheetDbClearMarks_(storage, row.row);
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
  storage.table.getRange(row, sheetDbConfig_().headers.length).setValue('');
  const cells = storage.table.getRange(row, 1, 1, sheetDbConfig_().headers.length - 1);
  if (cells.setBackground) cells.setBackground(null);
  if (cells.setNote) cells.setNote('');
}

function sheetDbReceiptValues_(receipt) {
  return sheetDbConfig_().opHeaders.map(function (key) { return receipt[key] === undefined ? '' : receipt[key]; });
}
function sheetDbReceipts_(storage) {
  return sheetDbCells_(storage.ops, 12).map(function (values, index) {
    const receipt = { storageRow: index + 3 };
    sheetDbConfig_().opHeaders.forEach(function (key, column) { receipt[key] = values[column]; });
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
    sheetDbAssert_(receipts.length < sheetDbConfig_().maxReceipts, 'LIMIT_EXCEEDED',
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
    if (receipt.status === 'APPLIED' && receipt.issuedAt <= Date.now() - sheetDbConfig_().receiptDuration)
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
      storage.table.getRange(candidate[0].row, sheetDbConfig_().headers.length - 1).setValue(receipt.targetRevision);
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
        storage.table.getRange(current[0].row, sheetDbConfig_().headers.length - 1).setValue(receipt.targetRevision); SpreadsheetApp.flush();
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
    sheetDbAssert_(state.rows.length < sheetDbConfig_().maxRecords, 'LIMIT_EXCEEDED', 'This table supports at most ' + sheetDbConfig_().maxRecords + ' current records.');
  } else {
    sheetDbAssert_(current, 'NOT_FOUND', 'This record does not exist.');
    sheetDbAssert_(current.revision === command.expectedRevision, 'REVISION_CONFLICT', 'The record changed. Read its current revision before editing.', { currentRevision: current.revision });
  }
  if (SHEET_DB_STATES_) sheetDbValidateMutation_(storage, state, command);
  sheetDbStable_(storage, state.rows);
  const receipt = {
    operationId: command.operationId, issuedAt: command.issuedAt, action: command.action, recordId: command.recordId,
    requestHash: signature, priorRevision: current ? current.revision : 0, priorHash: current ? current.hash : '',
    targetRevision: current ? current.revision + 1 : 1, targetHash: command.action === 'delete' ? '' : sheetDbDataHash_(SHEET_DB_CURRENT_ ? command.data : command.data.label),
    status: 'PREPARED', deleted: command.action === 'delete', row: current ? current.row : Math.max(3, storage.table.getLastRow() + 1)
  };
  // Capacity is known before dispatch; a full receipt table is not an uncertain write.
  sheetDbAssert_(sheetDbReceipts_(storage).length < sheetDbConfig_().maxReceipts, 'LIMIT_EXCEEDED',
    'The 24-hour retry receipt limit has been reached. Wait for old receipts to expire.');
  // Persist a payload-free intent before touching the record. Ambiguous interruption fails closed.
  let attemptedRecordWrite = false;
  try {
    sheetDbSaveReceipt_(storage, receipt);
    sheetDbStable_(storage, state.rows);
    attemptedRecordWrite = true;
    if (receipt.deleted) { storage.table.deleteRow(current.row); SpreadsheetApp.flush(); }
    else sheetDbWrite_(storage.table, receipt.row, sheetDbRowValues_(receipt.recordId, command.data, receipt.targetRevision));
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

// Generic app-declared schema. No application tables, hours or booking rules live here.
function sheetDbAllowed_(value, keys) {
  sheetDbAssert_(value && typeof value === 'object' && !Array.isArray(value) && Object.keys(value).every(function (key) { return keys.indexOf(key) !== -1; }),
    'VALIDATION_ERROR', 'Unsupported schema or record options.');
}
function sheetDbNormalizeSchema_(schema) {
  sheetDbAllowed_(schema, ['name', 'version', 'tables']);
  const name = function (value) { return typeof value === 'string' && /^[a-z][a-zA-Z0-9_]{0,31}$/.test(value) && ['records', 'constructor', 'prototype', 'id'].indexOf(value) === -1; };
  sheetDbAssert_(name(schema.name) && Number.isSafeInteger(schema.version) && schema.version > 0 && Array.isArray(schema.tables) && schema.tables.length > 0 && schema.tables.length <= 8,
    'VALIDATION_ERROR', 'Declare a named, versioned schema with one to eight tables.');
  const names = Object.create(null);
  const tables = schema.tables.map(function (table) {
    sheetDbAllowed_(table, ['name', 'fields', 'unique', 'maxRecords']);
    sheetDbAssert_(name(table.name) && table.name.toLowerCase() !== 'setup' && !names[table.name.toLowerCase()] && Array.isArray(table.fields) && table.fields.length > 0 && table.fields.length <= 12,
      'VALIDATION_ERROR', 'Table names must be distinct, cannot use the reserved Setup tab, and must have one to twelve fields.');
    names[table.name.toLowerCase()] = true;
    const fields = table.fields.map(function (field) {
      sheetDbAllowed_(field, ['name', 'type', 'required', 'min', 'max', 'choices', 'reference', 'when']);
      sheetDbAssert_(name(field.name) && ['text', 'integer', 'number', 'boolean', 'reference', 'date'].indexOf(field.type) !== -1,
        'VALIDATION_ERROR', 'Unsupported field name or type.');
      sheetDbAssert_(field.required === undefined || typeof field.required === 'boolean', 'VALIDATION_ERROR', 'required must be a boolean.');
      sheetDbAssert_(field.type !== 'reference' || typeof field.reference === 'string', 'VALIDATION_ERROR', 'Reference fields need a target table.');
      sheetDbAssert_(field.when === undefined || field.reference !== undefined, 'VALIDATION_ERROR', 'Field conditions apply to references.');
      const result = { name: field.name, type: field.type, required: field.required === true };
      ['min', 'max'].forEach(function (key) {
        if (field[key] !== undefined) {
          sheetDbAssert_(['integer', 'number'].indexOf(field.type) !== -1 && Number.isFinite(field[key]), 'VALIDATION_ERROR', 'Numeric bounds require numeric fields.');
          result[key] = field[key];
        }
      });
      sheetDbAssert_(result.min === undefined || result.max === undefined || result.min <= result.max, 'VALIDATION_ERROR', 'Invalid numeric bounds.');
      if (field.choices !== undefined) {
        sheetDbAssert_(Array.isArray(field.choices) && field.choices.length > 0 && field.choices.length <= 20 && field.choices.every(function (value) { return ['string', 'number', 'boolean'].indexOf(typeof value) !== -1; }), 'VALIDATION_ERROR', 'Invalid allowed values.');
        result.choices = field.choices;
      }
      if (field.reference !== undefined) { sheetDbAssert_(field.type === 'reference' && name(field.reference), 'VALIDATION_ERROR', 'Invalid reference.'); result.reference = field.reference; }
      if (field.when !== undefined) result.when = sheetDbPredicate_(field.when);
      return result;
    });
    sheetDbAssert_(new Set(fields.map(function (field) { return field.name; })).size === fields.length, 'VALIDATION_ERROR', 'Field names must be distinct.');
    sheetDbAssert_(table.unique === undefined || Array.isArray(table.unique) && table.unique.length <= 12, 'VALIDATION_ERROR', 'Invalid unique constraints.');
    const unique = (table.unique || []).map(function (rule) {
      sheetDbAllowed_(rule, ['fields', 'when']);
      sheetDbAssert_(Array.isArray(rule.fields) && rule.fields.length > 0 && rule.fields.every(function (field) { return fields.some(function (item) { return item.name === field; }); }), 'VALIDATION_ERROR', 'Invalid unique constraint.');
      return { fields: rule.fields, ...(rule.when ? { when: sheetDbPredicate_(rule.when) } : {}) };
    });
    const maxRecords = table.maxRecords === undefined ? 1000 : table.maxRecords;
    sheetDbAssert_(Number.isSafeInteger(maxRecords) && maxRecords >= 1 && maxRecords <= 1000, 'VALIDATION_ERROR', 'Invalid record limit.');
    fields.concat(unique).forEach(function (rule) {
      if (rule.when) sheetDbAssert_(fields.some(function (field) { return field.name === rule.when.field; }), 'VALIDATION_ERROR', 'Unknown condition field.');
    });
    return { name: table.name, fields: fields, unique: unique, maxRecords: maxRecords };
  });
  tables.forEach(function (table) { table.fields.forEach(function (field) {
    if (field.reference) sheetDbAssert_(tables.some(function (target) { return target.name === field.reference; }), 'VALIDATION_ERROR', 'Reference target is undeclared.');
  }); });
  const result = { name: schema.name, version: schema.version, tables: tables };
  sheetDbAssert_(encodeURIComponent(JSON.stringify(result)).replace(/%[A-F0-9]{2}|./g, 'x').length <= 8000, 'LIMIT_EXCEEDED', 'The schema exceeds this release’s configuration limit.');
  return result;
}
function sheetDbPredicate_(condition) {
  sheetDbAllowed_(condition, ['field', 'equals']);
  sheetDbAssert_(typeof condition.field === 'string' && ['string', 'number', 'boolean'].indexOf(typeof condition.equals) !== -1, 'VALIDATION_ERROR', 'Invalid condition.');
  return { field: condition.field, equals: condition.equals };
}
function sheetDbLoadSchema_() {
  const text = PropertiesService.getDocumentProperties().getProperty('sheetdb.applicationSchema.v1');
  if (!text) return null;
  try { return sheetDbNormalizeSchema_(JSON.parse(text)); }
  catch (_) { sheetDbFail_('SCHEMA_CHANGED', 'The application schema is unreadable. Restore its configuration before continuing.'); }
}
function sheetDbTableConfig_(table) {
  return Object.assign({}, SHEET_DB_PROOF_, {
    sheetName: table.name, metaName: '_sheetdb_' + table.name + '_meta', opsName: '_sheetdb_' + table.name + '_ops',
    initializedKey: 'sheetdb.initialized.v2.' + table.name, initializingKey: 'sheetdb.initializing.v2.' + table.name,
    headers: ['id'].concat(table.fields.map(function (field) { return field.name; }), ['_revision', '_status']),
    types: ['id'].concat(table.fields.map(function (field) { return field.type; }), ['integer', 'status']),
    fields: table.fields, unique: table.unique, maxRecords: table.maxRecords
  });
}
function sheetDbValidateData_(input) {
  sheetDbAssert_(input && typeof input === 'object' && !Array.isArray(input), 'VALIDATION_ERROR', 'A complete record is required.');
  const fields = sheetDbFields_(), data = {};
  sheetDbAssert_(Object.keys(input).every(function (key) { return fields.some(function (field) { return field.name === key; }); }), 'VALIDATION_ERROR', 'The record contains unknown fields.');
  fields.forEach(function (field) {
    const value = input[field.name] === undefined ? '' : input[field.name];
    if (value === '' && !field.required) { data[field.name] = ''; return; }
    let valid;
    if (field.type === 'integer') valid = Number.isSafeInteger(value);
    else if (field.type === 'number') valid = typeof value === 'number' && Number.isFinite(value);
    else if (field.type === 'boolean') valid = typeof value === 'boolean';
    else valid = typeof value === 'string' && value.trim().length > 0 && value.length <= 500 && !value.startsWith('=');
    if (valid && field.type === 'date') valid = /^\d{4}-\d{2}-\d{2}$/.test(value) && !isNaN(Date.parse(value + 'T00:00:00Z')) && new Date(value + 'T00:00:00Z').toISOString().slice(0, 10) === value;
    if (valid && field.min !== undefined) valid = value >= field.min;
    if (valid && field.max !== undefined) valid = value <= field.max;
    if (valid && field.choices) valid = field.choices.indexOf(value) !== -1;
    sheetDbAssert_(valid, 'VALIDATION_ERROR', 'Invalid value for ' + field.name + ' (' + field.type + ').', { field: field.name });
    data[field.name] = value;
  });
  return data;
}
function sheetDbCondition_(data, condition) { return !condition || data[condition.field] === condition.equals; }
function sheetDbCheckTables_(states, mark) {
  const issues = [];
  Object.keys(states).forEach(function (name) {
    const state = states[name]; SHEET_DB_CURRENT_ = state.config;
    state.rows.forEach(function (row) {
      try {
        sheetDbAssert_(!row.formula, 'VALIDATION_ERROR', 'Formulas are not supported.');
        sheetDbValidateData_(row.data);
        state.config.fields.forEach(function (field) {
          if (!field.reference || !sheetDbCondition_(row.data, field.when) || row.data[field.name] === '') return;
          sheetDbAssert_(states[field.reference].rows.some(function (target) { return target.id === row.data[field.name]; }), 'REFERENCE_CONFLICT', 'Referenced record does not exist: ' + field.name + '.');
        });
      } catch (error) { issues.push({ table: name, row: row.row, field: 'record', message: error.message }); }
    });
    state.config.unique.forEach(function (rule) {
      const seen = Object.create(null);
      state.rows.forEach(function (row) {
        if (!sheetDbCondition_(row.data, rule.when)) return;
        const key = JSON.stringify(rule.fields.map(function (field) { return row.data[field]; }));
        if (seen[key]) issues.push({ table: name, row: row.row, field: rule.fields.join(', '), message: 'A unique value is already in use.' });
        seen[key] = true;
      });
    });
  });
  if (mark && issues.length) issues.forEach(function (issue) {
    const state = states[issue.table]; SHEET_DB_CURRENT_ = state.config;
    try { state.storage.table.getRange(issue.row, state.config.headers.length).setValue('Needs attention'); } catch (_) {}
  });
  sheetDbAssert_(!issues.length, mark ? 'MANUAL_EDIT_INVALID' : 'CONSTRAINT_VIOLATION', 'The requested change violates the declared table rules.', { issues: issues });
}
function sheetDbValidateMutation_(storage, state, command) {
  const targetConfig = sheetDbConfig_(), candidate = Object.assign({}, SHEET_DB_STATES_);
  let rows = state.rows.filter(function (row) { return row.id !== command.recordId; });
  if (command.action !== 'delete') rows = rows.concat([{ id: command.recordId, data: command.data, label: command.data, revision: command.expectedRevision + 1 || 1, row: state.byId[command.recordId] ? state.byId[command.recordId].row : Math.max(3, storage.table.getLastRow() + 1), formula: false }]);
  candidate[targetConfig.sheetName] = Object.assign({}, candidate[targetConfig.sheetName], { rows: rows });
  try {
    Object.keys(SHEET_DB_STATES_).forEach(function (name) {
      SHEET_DB_CURRENT_ = SHEET_DB_STATES_[name].config;
      sheetDbStable_(SHEET_DB_STATES_[name].storage, SHEET_DB_STATES_[name].rows);
    });
    sheetDbCheckTables_(candidate, false);
  } finally { SHEET_DB_CURRENT_ = targetConfig; }
}
function sheetDbTableRequest_(input) {
  const lock = LockService.getDocumentLock();
  sheetDbAssert_(lock && lock.tryLock(5000), 'BUSY', 'The database is busy. Retry the same request.');
  try {
    let schema = sheetDbLoadSchema_();
    const ensure = input.action === 'ensureSchema';
    if (ensure) {
      sheetDbAllowed_(input, ['action', 'schema']);
      const desired = sheetDbNormalizeSchema_(input.schema);
      sheetDbAssert_(!schema || JSON.stringify(schema) === JSON.stringify(desired), 'SCHEMA_MISMATCH', 'This database has a different schema. Automatic schema migrations are not supported.');
      if (!schema) PropertiesService.getDocumentProperties().setProperty('sheetdb.applicationSchema.v1', JSON.stringify(desired));
      schema = desired;
    }
    sheetDbAssert_(schema, 'SCHEMA_REQUIRED', 'The app must declare its schema before accessing tables.');
    const states = Object.create(null);
    schema.tables.forEach(function (table) {
      SHEET_DB_CURRENT_ = sheetDbTableConfig_(table);
      const storage = sheetDbStorage_(); sheetDbRecover_(storage); sheetDbPrune_(storage);
      states[table.name] = { config: SHEET_DB_CURRENT_, storage: storage, rows: sheetDbReadRows_(storage) };
    });
    sheetDbCheckTables_(states, true);
    schema.tables.forEach(function (table) {
      SHEET_DB_CURRENT_ = states[table.name].config;
      Object.assign(states[table.name], sheetDbReconcile_(states[table.name].storage));
    });
    sheetDbCheckTables_(states, true);
    if (ensure) return { result: { name: schema.name, version: schema.version, tables: schema.tables.map(function (table) { return table.name; }) } };
    if (input.action === 'snapshot') return { result: Object.fromEntries(Object.keys(states).map(function (name) { return [name, states[name].rows.map(sheetDbRecord_)]; })) };
    sheetDbAssert_(Object.prototype.hasOwnProperty.call(states, input.table), 'TABLE_NOT_FOUND', 'This table has not been declared by the app.');
    const state = states[input.table]; SHEET_DB_CURRENT_ = state.config; SHEET_DB_STATES_ = states;
    const raw = Object.assign({}, input); delete raw.table;
    const command = sheetDbCommand_(raw); command.table = input.table;
    if (command.action === 'list') return { result: state.rows.map(sheetDbRecord_) };
    if (command.action === 'read') {
      sheetDbAssert_(state.byId[command.recordId], 'NOT_FOUND', 'This record does not exist.');
      return { result: sheetDbRecord_(state.byId[command.recordId]) };
    }
    return { result: sheetDbMutate_(state.storage, state, command) };
  } finally { SHEET_DB_CURRENT_ = null; SHEET_DB_STATES_ = null; lock.releaseLock(); }
}
