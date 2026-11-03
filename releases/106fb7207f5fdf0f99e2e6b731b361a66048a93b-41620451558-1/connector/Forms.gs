/**
 * Public Google Forms transport POC, protocol 1. Deliberately NO write authentication.
 * Use disposable data: anyone who knows the Form can submit database commands.
 * No web-app deployment, endpoint or open spreadsheet tab is needed after setup.
 */
const SHEET_DB_FORMS_ = Object.freeze({
  property: 'sheetdb.forms.poc.v1', responseName: '_sheetdb_public',
  commandTitle: 'sheet-db command', handler: 'sheetDbFormsSubmit',
  headers: ['requestId', 'response', 'createdAt', 'requestHash'],
  maxResponses: 1000, maxEnvelope: 45000, maxResponse: 45000
});

function sheetDbFormsState_(book) {
  let state;
  try { state = JSON.parse(PropertiesService.getDocumentProperties().getProperty(SHEET_DB_FORMS_.property) || 'null'); }
  catch (_) { state = null; }
  // A copied template must create resources for its own workbook, never the original's Form.
  return state && state.databaseId === book.getId() ? state : { databaseId: book.getId() };
}
function sheetDbFormsSave_(state) {
  PropertiesService.getDocumentProperties().setProperty(SHEET_DB_FORMS_.property, JSON.stringify(state));
}
function sheetDbFormsResponses_(book) {
  let sheet = book.getSheetByName(SHEET_DB_FORMS_.responseName);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_DB_FORMS_.responseName);
    sheet.getRange(1, 1, 1, 4).setValues([SHEET_DB_FORMS_.headers]);
    sheet.setFrozenRows(1);
  }
  sheetDbHeaders_(sheet, [SHEET_DB_FORMS_.headers], 4, 'SCHEMA_CHANGED');
  return sheet;
}
function sheetDbFormsConfig_(book, state) {
  const sheet = sheetDbFormsResponses_(book);
  sheetDbAssert_(state.formUrl && state.commandEntry && Number.isInteger(state.inboxSheetId),
    'SETUP_REQUIRED', 'Choose sheet-db → Set up public Forms POC first.');
  return {
    protocol: 1, spreadsheetId: book.getId(), spreadsheetName: book.getName(),
    formUrl: state.formUrl, commandEntry: state.commandEntry, responseSheetId: sheet.getSheetId()
  };
}

/** Template owners install this once; each copied database runs it under its new owner. */
function sheetDbSetupFormsPoc() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Public Forms proof of concept',
    'This makes this entire spreadsheet viewable by anyone with its link. Its Form allows unauthenticated writes. ' +
    'Form submissions and operation responses retain submitted data. Use disposable test data only. Continue?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const book = SpreadsheetApp.getActiveSpreadsheet(), confirmedId = book && book.getId();
  const outcome = sheetDbSafe_(function () {
    return { result: sheetDbSetupFormsPoc_(book, function () { return book && book.getId() === confirmedId; }) };
  });
  if (outcome.error) { ui.alert(outcome.error.code + ': ' + outcome.error.message); return outcome; }
  sheetDbFormsShow_(outcome.result);
  return outcome;
}

function sheetDbSetupFormsPoc_(book, confirmPublicAccess) {
  book = book || SpreadsheetApp.getActiveSpreadsheet();
  sheetDbAssert_(book, 'SETUP_REQUIRED', 'Run setup from the database spreadsheet.');
  const confirmed = function () {
    sheetDbAssert_(typeof confirmPublicAccess === 'function' && confirmPublicAccess() === true,
      'PUBLIC_CONFIRMATION_REQUIRED', 'Confirm public access for this database before running setup.');
  };
  confirmed();
  const lock = LockService.getScriptLock();
  sheetDbAssert_(lock.tryLock(5000), 'BUSY', 'Setup is already running. Wait, then retry.');
  const previousBook = SHEET_DB_BOOK_;
  try {
    SHEET_DB_BOOK_ = book;
    confirmed();
    const state = sheetDbFormsState_(book);
    const response = sheetDbFormsResponses_(book);
    // Initialize the normal readable records table without creating a panel session.
    const initialized = sheetDbSafe_(function () { return sheetDbExecuteCommand_({ action: 'list' }); });
    sheetDbAssert_(!initialized.error, initialized.error && initialized.error.code || 'SETUP_FAILED',
      initialized.error && initialized.error.message || 'The records table could not be initialized.');

    let form;
    if (state.formId) form = FormApp.openById(state.formId);
    else {
      form = FormApp.create('sheet-db POC · ' + book.getName(), true);
      state.formId = form.getId();
      sheetDbFormsSave_(state); // Save immediately so a retry reuses an already-created Form.
    }
    form.setDescription('Public, unauthenticated sheet-db transport proof. Use disposable data only.')
      .setCollectEmail(false).setLimitOneResponsePerUser(false).setAllowResponseEdits(false)
      .setPublishingSummary(false).setShowLinkToRespondAgain(false);
    if (form.supportsAdvancedResponderPermissions()) form.setPublished(true);
    form.setAcceptingResponses(true);

    let item = form.getItems().filter(function (candidate) { return candidate.getTitle() === SHEET_DB_FORMS_.commandTitle; })[0];
    if (item) item = item.asParagraphTextItem();
    else item = form.addParagraphTextItem().setTitle(SHEET_DB_FORMS_.commandTitle).setRequired(true);
    const prefilled = form.createResponse().withItemResponse(item.createResponse('sheetdb-bootstrap')).toPrefilledUrl();
    const field = /[?&](entry\.\d+)=/.exec(prefilled);
    sheetDbAssert_(field, 'SETUP_FAILED', 'Google did not provide a Form field identifier.');
    state.commandEntry = field[1];
    state.formUrl = form.getPublishedUrl().replace(/\/viewform(?:\?.*)?$/, '/formResponse');
    sheetDbAssert_(/\/formResponse$/.test(state.formUrl), 'SETUP_FAILED', 'Google returned an unexpected Form address.');
    sheetDbFormsSave_(state);

    let destinationId = null;
    try { destinationId = form.getDestinationId(); }
    catch (error) {
      // Google throws for a new, unlinked Form instead of returning null.
      // Permission/service failures must still stop setup without changing its destination.
      if (!/no response destination/i.test(String(error && error.message || error))) throw error;
    }
    if (destinationId !== book.getId()) form.setDestination(FormApp.DestinationType.SPREADSHEET, book.getId());
    SpreadsheetApp.flush();
    const inbox = book.getSheets().filter(function (sheet) {
      // Copies can retain a link to an unrelated Form the new owner cannot edit.
      // Such a link must not prevent discovering the Form created for this copy.
      try {
        const url = sheet.getFormUrl();
        return url && FormApp.openByUrl(url).getId() === state.formId;
      } catch (_) { return false; }
    })[0];
    sheetDbAssert_(inbox, 'SETUP_FAILED', 'Google has not created the Form response sheet yet. Retry setup.');
    state.inboxSheetId = inbox.getSheetId();
    state.responseSheetId = response.getSheetId();
    sheetDbFormsSave_(state);

    const matching = ScriptApp.getProjectTriggers().filter(function (trigger) {
      return trigger.getHandlerFunction() === SHEET_DB_FORMS_.handler && trigger.getTriggerSourceId() === book.getId();
    });
    if (!matching.length) ScriptApp.newTrigger(SHEET_DB_FORMS_.handler).forSpreadsheet(book).onFormSubmit().create();
    matching.slice(1).forEach(function (trigger) { ScriptApp.deleteTrigger(trigger); });

    // Sharing the Form permits respondents, not script editing. Account policy may reject this.
    confirmed();
    DriveApp.getFileById(state.formId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    confirmed();
    DriveApp.getFileById(book.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    state.ready = true;
    sheetDbFormsSave_(state);
    return sheetDbFormsConfig_(book, state);
  } finally { SHEET_DB_BOOK_ = previousBook; lock.releaseLock(); }
}

function sheetDbShowFormsConnection() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const outcome = sheetDbSafe_(function () {
    const state = sheetDbFormsState_(book);
    sheetDbAssert_(state.ready, 'SETUP_REQUIRED', 'Choose sheet-db → Set up public Forms POC first.');
    return { result: sheetDbFormsConfig_(book, state) };
  });
  if (outcome.error) { SpreadsheetApp.getUi().alert(outcome.error.code + ': ' + outcome.error.message); return outcome; }
  sheetDbFormsShow_(outcome.result);
  return outcome;
}
function sheetDbFormsShow_(config) {
  const json = JSON.stringify(config, null, 2).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const html = '<!doctype html><html><body style="font:14px system-ui;padding:16px">' +
    '<h2>Connect your app</h2><p>Copy these connection details into the app once. Afterwards you can close this spreadsheet.</p>' +
    '<textarea readonly aria-label="Connection details" onclick="this.select()" style="width:100%;height:230px">' + json + '</textarea>' +
    '<p><strong>Public, unauthenticated POC.</strong> Anyone with the Form can change the database. Use disposable data.</p></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(430), 'Public Forms connection');
}

/** Installable spreadsheet Form-submit trigger. No browser session or active Sheet tab is needed. */
function sheetDbFormsSubmit(event) {
  if (!event || !event.source || !event.range || !Array.isArray(event.values)) return;
  const book = event.source, state = sheetDbFormsState_(book);
  if (!state.ready || state.databaseId !== book.getId() || event.range.getSheet().getSheetId() !== state.inboxSheetId) return;
  const raw = event.values[1];
  if (typeof raw !== 'string' || raw.length > SHEET_DB_FORMS_.maxEnvelope) return;
  let envelope;
  try { envelope = JSON.parse(raw); } catch (_) { return; }
  if (!envelope || typeof envelope !== 'object' || Array.isArray(envelope) ||
      typeof envelope.requestId !== 'string' || !/^[A-Za-z0-9_-]{16,128}$/.test(envelope.requestId)) return;
  const lock = LockService.getScriptLock();
  // If delivery/quotas prevent handling, the app reports an unconfirmed request and may retry.
  lock.waitLock(20000);
  try {
    const sheet = sheetDbFormsResponses_(book), requestHash = sheetDbHash_(raw);
    const existing = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().filter(function (row) {
      return row[0] === envelope.requestId;
    })[0] : null;
    if (existing) {
      // Never replace a completed request result; mutations have separate durable operation IDs.
      // A conflicting duplicate is ignored, so it cannot overwrite an original acknowledgement.
      if (existing[3] !== requestHash) console.warn('Ignored conflicting sheet-db request ID.');
      return;
    }
    const previousBook = SHEET_DB_BOOK_;
    SHEET_DB_BOOK_ = book;
    let outcome;
    try {
      outcome = sheetDbSafe_(function () {
        sheetDbAllowed_(envelope, ['protocol', 'databaseId', 'requestId', 'command']);
        sheetDbAssert_(envelope.protocol === 1, 'PROTOCOL_MISMATCH', 'Unsupported public Forms protocol.');
        sheetDbAssert_(envelope.databaseId === book.getId(), 'IDENTITY_MISMATCH', 'This command belongs to a different database.');
        sheetDbAssert_(envelope.command && typeof envelope.command === 'object' && !Array.isArray(envelope.command), 'VALIDATION_ERROR', 'A database command is required.');
        sheetDbAssert_(['create', 'read', 'list', 'update', 'delete', 'ensureSchema', 'snapshot'].indexOf(envelope.command.action) !== -1,
          'VALIDATION_ERROR', 'Unsupported database command.');
        return sheetDbExecuteCommand_(envelope.command);
      });
    } finally { SHEET_DB_BOOK_ = previousBook; }
    const response = Object.assign({ protocol: 1, databaseId: book.getId(), requestId: envelope.requestId }, outcome);
    let encoded = JSON.stringify(response);
    if (encoded.length > SHEET_DB_FORMS_.maxResponse) {
      encoded = JSON.stringify({ protocol: 1, databaseId: book.getId(), requestId: envelope.requestId,
        error: { code: 'RESPONSE_TOO_LARGE', message: 'This proof supports response JSON up to 45,000 characters. Use fewer records.' } });
    }
    if (sheet.getLastRow() > SHEET_DB_FORMS_.maxResponses) sheet.deleteRows(2, sheet.getLastRow() - SHEET_DB_FORMS_.maxResponses);
    const row = sheet.getLastRow() + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), row - sheet.getMaxRows());
    sheet.getRange(row, 1, 1, 4).setNumberFormat('@').setValues([[envelope.requestId, encoded, new Date().toISOString(), requestHash]]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
