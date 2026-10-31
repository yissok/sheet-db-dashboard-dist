/**
 * sheet-db 0.1.0 — private, bound-Sheet transport proof.
 * @OnlyCurrentDoc
 *
 * No web-app deployment, HTTP endpoint, token export, or third-party service.
 * The single public RPC is called by this project's HTML-service panel.
 */
const SHEET_DB_PROOF_ = Object.freeze({
  protocol: 1,
  schema: 1,
  version: '0.1.0',
  sheetName: '_sheetdb_proof_v1',
  header: 'sheet-db proof journal v1',
  sessionKey: 'sheetdb.proof.connection.v1',
  sessionDuration: 30 * 60 * 1000,
  maxMutations: 1000
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('sheet-db')
    .addItem('Open connection panel', 'openSheetDbPanel').addToUi();
}

function openSheetDbPanel() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Panel').setTitle('sheet-db connection')
  );
}

/** Authorize one short-lived connection for the current Google account. */
function sheetDbPrepareConnection(input) {
  return sheetDbSafe_(function () {
    sheetDbAssert_(input && typeof input === 'object', 'VALIDATION_ERROR', 'Connection information is missing.');
    sheetDbAssert_(typeof input.channel === 'string' && /^[A-Za-z0-9_-]{43}$/.test(input.channel),
      'VALIDATION_ERROR', 'The connection code is invalid.');
    sheetDbAssert_(typeof input.appOrigin === 'string' && sheetDbValidOrigin_(input.appOrigin),
      'VALIDATION_ERROR', 'Use an HTTPS app address, or localhost for development.');
    const now = Date.now();
    sheetDbAssert_(Number.isFinite(input.expiresAt) && input.expiresAt > now && input.expiresAt <= now + 5 * 60 * 1000,
      'CONNECTION_EXPIRED', 'This connection invitation has expired. Open the app again.');
    const identity = sheetDbIdentity_();
    const sessionExpiresAt = now + SHEET_DB_PROOF_.sessionDuration;
    PropertiesService.getUserProperties().setProperty(SHEET_DB_PROOF_.sessionKey, JSON.stringify({
      channel: input.channel,
      appOrigin: input.appOrigin,
      accountId: identity.accountId,
      databaseId: identity.databaseId,
      expiresAt: sessionExpiresAt
    }));
    return { identity: identity, sessionExpiresAt: sessionExpiresAt };
  });
}

/** JSON-only RPC. Every request is bound to the prepared account and Sheet. */
function sheetDbExecute(request) {
  return sheetDbSafe_(function () {
    sheetDbAssert_(request && typeof request === 'object', 'VALIDATION_ERROR', 'The request is missing.');
    sheetDbAssert_(request.protocol === SHEET_DB_PROOF_.protocol,
      'PROTOCOL_MISMATCH', 'Update the app and connector to compatible versions.');
    const identity = sheetDbIdentity_();
    const encoded = PropertiesService.getUserProperties().getProperty(SHEET_DB_PROOF_.sessionKey);
    let session;
    try { session = encoded ? JSON.parse(encoded) : null; } catch (_) { session = null; }
    sheetDbAssert_(session && session.channel === request.channel && session.expiresAt > Date.now(),
      'CONNECTION_EXPIRED', 'Open the connection panel and connect the app again.');
    sheetDbAssert_(session.accountId === identity.accountId && session.databaseId === identity.databaseId &&
      request.context && request.context.accountId === identity.accountId && request.context.databaseId === identity.databaseId,
      'IDENTITY_MISMATCH', 'The account or database changed. Reconnect before trying again.');
    const command = sheetDbCommand_(request.command);
    const lock = LockService.getDocumentLock();
    sheetDbAssert_(lock && lock.tryLock(5000), 'BUSY', 'The database is busy. Try the same request again.', { retryable: true });
    try {
      const sheet = sheetDbJournalSheet_();
      const state = sheetDbReplay_(sheet);
      if (command.action === 'list') {
        return { result: Object.keys(state.records).map(function (id) { return state.records[id]; })
          .filter(function (record) { return !record.deleted; }) };
      }
      if (command.action === 'read') {
        const record = state.records[command.recordId];
        sheetDbAssert_(record && !record.deleted, 'NOT_FOUND', 'This record does not exist.');
        return { result: record };
      }
      const previous = state.operations[command.operationId];
      if (previous) {
        sheetDbAssert_(JSON.stringify(previous.command) === JSON.stringify(command),
          'IDEMPOTENCY_CONFLICT', 'This operation code was already used for a different change.');
        return { result: previous.result };
      }
      sheetDbAssert_(state.count < SHEET_DB_PROOF_.maxMutations, 'LIMIT_EXCEEDED',
        'This proof database has reached its 1,000-change limit. Keep it for verification and use a fresh test Sheet.');
      const result = sheetDbApply_(state.records, command);
      const entry = { schemaVersion: SHEET_DB_PROOF_.schema, operationId: command.operationId, command: command, result: result };
      const serialized = JSON.stringify(entry);
      const nextRow = state.count + 2;
      try {
        const maxRows = sheet.getMaxRows();
        if (nextRow > maxRows) sheet.insertRowsAfter(maxRows, nextRow - maxRows);
        sheet.getRange(nextRow, 1).setValue(serialized);
        SpreadsheetApp.flush();
        const readBack = sheet.getRange(nextRow, 1).getValues()[0][0];
        sheetDbAssert_(readBack === serialized, 'WRITE_UNCONFIRMED', 'The write could not be verified.');
      } catch (_) {
        sheetDbFail_('WRITE_UNCONFIRMED', 'The write may have completed. Retry with the same operation code to check it.',
          { operationId: command.operationId, retryable: true });
      }
      return { result: result };
    } finally {
      lock.releaseLock();
    }
  });
}

function sheetDbIdentity_() {
  const effective = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const active = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  sheetDbAssert_(effective && /^[^\s@]+@[^\s@]+$/.test(effective), 'IDENTITY_UNAVAILABLE',
    'Google did not identify the authorized account. Use a browser profile signed into only the Sheet owner and try again.');
  sheetDbAssert_(!active || active === effective, 'IDENTITY_MISMATCH',
    'Google returned different active and authorized accounts. Use a browser profile signed into only the Sheet owner.');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheetDbAssert_(spreadsheet, 'IDENTITY_UNAVAILABLE', 'Open this connector from its own Google Sheet.');
  return {
    accountId: effective,
    accountEmail: effective,
    databaseId: spreadsheet.getId(),
    databaseName: spreadsheet.getName(),
    protocolVersion: SHEET_DB_PROOF_.protocol,
    schemaVersion: SHEET_DB_PROOF_.schema,
    connectorVersion: SHEET_DB_PROOF_.version
  };
}

function sheetDbCommand_(input) {
  sheetDbAssert_(input && typeof input === 'object' && !Array.isArray(input), 'VALIDATION_ERROR', 'The command is invalid.');
  const action = input.action;
  sheetDbAssert_(['create', 'read', 'update', 'delete', 'list'].indexOf(action) !== -1,
    'VALIDATION_ERROR', 'This command is not supported by the proof.');
  const allowed = { action: true };
  const normalized = { action: action };
  if (action === 'create' || action === 'update' || action === 'delete') {
    allowed.operationId = true;
    sheetDbAssert_(typeof input.operationId === 'string' && /^[A-Za-z0-9_-]{16,128}$/.test(input.operationId),
      'VALIDATION_ERROR', 'A stable operation code is required for every change.');
    normalized.operationId = input.operationId;
  }
  if (action === 'create' || action === 'read' || action === 'update' || action === 'delete') {
    allowed.recordId = true;
    sheetDbAssert_(typeof input.recordId === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(input.recordId),
      'VALIDATION_ERROR', 'The record identifier is invalid.');
    normalized.recordId = input.recordId;
  }
  if (action === 'update' || action === 'delete') {
    allowed.expectedRevision = true;
    sheetDbAssert_(Number.isSafeInteger(input.expectedRevision) && input.expectedRevision >= 1,
      'VALIDATION_ERROR', 'The record version is required before changing it.');
    normalized.expectedRevision = input.expectedRevision;
  }
  if (action === 'create' || action === 'update') {
    allowed.data = true;
    sheetDbAssert_(input.data && typeof input.data === 'object' && !Array.isArray(input.data) &&
      Object.keys(input.data).length === 1 && Object.prototype.hasOwnProperty.call(input.data, 'label') &&
      typeof input.data.label === 'string' && input.data.label.length >= 1 && input.data.label.length <= 500 && input.data.label.trim(),
      'VALIDATION_ERROR', 'Enter a label between 1 and 500 characters.');
    normalized.data = { label: input.data.label };
  }
  sheetDbAssert_(Object.keys(input).every(function (key) { return Object.prototype.hasOwnProperty.call(allowed, key); }),
    'VALIDATION_ERROR', 'The command contains unsupported fields.');
  return normalized;
}

function sheetDbJournalSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_DB_PROOF_.sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_DB_PROOF_.sheetName);
    sheet.getRange(1, 1).setValue(SHEET_DB_PROOF_.header);
    SpreadsheetApp.flush();
  }
  sheetDbAssert_(sheet.getLastRow() >= 1 && sheet.getRange(1, 1).getValues()[0][0] === SHEET_DB_PROOF_.header,
    'STORAGE_CORRUPT', 'The test storage was edited or is incomplete. Keep this Sheet for inspection and use a fresh test copy.');
  return sheet;
}

function sheetDbReplay_(sheet) {
  const state = { records: Object.create(null), operations: Object.create(null), count: sheet.getLastRow() - 1 };
  sheetDbAssert_(state.count <= SHEET_DB_PROOF_.maxMutations, 'STORAGE_CORRUPT', 'The proof journal exceeds its supported size.');
  const rows = state.count ? sheet.getRange(2, 1, state.count, 1).getValues() : [];
  rows.forEach(function (row) {
    let entry;
    try { entry = JSON.parse(row[0]); } catch (_) { sheetDbFail_('STORAGE_CORRUPT', 'The test storage contains an unreadable change.'); }
    try {
      sheetDbAssert_(entry && entry.schemaVersion === SHEET_DB_PROOF_.schema && entry.command &&
        ['create', 'update', 'delete'].indexOf(entry.command.action) !== -1 &&
        entry.operationId === entry.command.operationId && !state.operations[entry.operationId],
        'STORAGE_CORRUPT', 'The test storage contains an invalid change.');
      const command = sheetDbCommand_(entry.command);
      // Verify the persisted client ID; replay must not invent IDs.
      const expected = sheetDbApply_(state.records, command, entry.result && entry.result.id);
      sheetDbAssert_(JSON.stringify(expected) === JSON.stringify(entry.result), 'STORAGE_CORRUPT', 'The test storage has conflicting changes.');
      state.records[entry.result.id] = entry.result;
      state.operations[entry.operationId] = entry;
    } catch (_) {
      sheetDbFail_('STORAGE_CORRUPT', 'The test storage was edited or has conflicting changes. Keep it for inspection.');
    }
  });
  return state;
}

function sheetDbApply_(records, command, persistedId) {
  if (command.action === 'create') {
    const id = command.recordId;
    sheetDbAssert_(!persistedId || persistedId === id, 'STORAGE_CORRUPT', 'A stored record identifier was changed.');
    sheetDbAssert_(typeof id === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(id),
      'STORAGE_CORRUPT', 'The stored record identifier is invalid.');
    sheetDbAssert_(!records[id], 'ALREADY_EXISTS', 'This record identifier already exists.');
    return { id: id, revision: 1, data: { label: command.data.label }, deleted: false };
  }
  const current = records[command.recordId];
  sheetDbAssert_(current && !current.deleted, 'NOT_FOUND', 'This record does not exist.');
  sheetDbAssert_(current.revision === command.expectedRevision, 'REVISION_CONFLICT',
    'This record changed since it was loaded. Read its current version before editing.', { currentRevision: current.revision });
  return {
    id: current.id,
    revision: current.revision + 1,
    data: command.action === 'delete' ? current.data : { label: command.data.label },
    deleted: command.action === 'delete'
  };
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
    const message = String(error && error.message ? error.message : 'Google could not complete this request.').slice(0, 400);
    const quota = /quota|too many times|limit exceeded/i.test(message);
    return { error: { code: quota ? 'QUOTA_EXCEEDED' : 'GOOGLE_ERROR', message: message, details: { retryable: false } } };
  }
}

function sheetDbAssert_(condition, code, message, details) {
  if (!condition) sheetDbFail_(code, message, details);
}

function sheetDbFail_(code, message, details) {
  const error = new Error(message);
  error.sheetDbCode = code;
  error.sheetDbDetails = details;
  throw error;
}
