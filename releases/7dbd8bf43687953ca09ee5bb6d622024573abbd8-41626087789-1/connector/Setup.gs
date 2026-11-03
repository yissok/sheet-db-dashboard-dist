/** Shared desktop/mobile setup. Browser callbacks require no web-app deployment. */
const SHEET_DB_SETUP_ = Object.freeze({
  sheetName: 'Setup', marker: 'sheet-db setup v2', property: 'sheetdb.setup.v1',
  handler: 'sheetDbSetupEdit', lifetimeSeconds: 900,
  disclosure: 'This entire spreadsheet will be viewable by anyone with its link. Its Form permits unauthenticated writes. ' +
    'Form submissions and operation responses retain submitted data. Use disposable test data only.',
  apps: Object.freeze({
    timeslot: 'https://scout-38aa2.web.app/timeslot/',
    dashboard: 'https://scout-38aa2.web.app/sheet-db-dashboard/'
  })
});

/** Maintainer: prepare, upgrade managed v1, or finish formatting managed v2. */
function sheetDbPrepareSetupTemplate() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  sheetDbAssert_(book, 'SETUP_REQUIRED', 'Open the database spreadsheet first.');
  let sheet = book.getSheetByName(SHEET_DB_SETUP_.sheetName), previousConfig, completed;
  if (sheet && sheet.getRange(1, 1).getValue() === SHEET_DB_SETUP_.marker) {
    // A previous run may have written the layout before formatting failed.
    // Existing controls, consent and connection outputs must stay unchanged.
    sheetDbSetupFormatTemplate_(sheet);
    if (!sheetDbSetupOwns_(sheetDbSetupRead_(), book, sheet)) {
      sheetDbSetupSave_({ databaseId: book.getId(), sheetId: sheet.getSheetId() });
    }
    book.setActiveSheet(sheet);
    return;
  }
  if (sheet) {
    sheetDbAssert_(sheet.getRange(1, 1).getValue() === 'sheet-db setup v1', 'SETUP_EXISTS',
      'Setup already exists. Only the managed v1 layout can be upgraded; other sheets are left unchanged.');
    sheetDbAssert_(!sheet.getRange(1, 4, 16, 2).getValues().some(function (row) {
      return row.some(function (value) { return value !== ''; });
    }), 'SETUP_EXISTS', 'Columns D–E contain data. Move it before upgrading Setup.');
    try { previousConfig = JSON.parse(sheet.getRange(9, 2).getValue()); } catch (_) {}
    completed = sheet.getRange(16, 2).getValue();
    sheet.getRange(1, 1, 16, 2).clearContent().clearDataValidations().clearFormat();
  } else sheet = book.insertSheet(SHEET_DB_SETUP_.sheetName);
  sheet.getRange(1, 1, 7, 2).setValues([
    [SHEET_DB_SETUP_.marker, 'Follow the app guide. Copy URLs into your private browser.'],
    ['1. Prepare links', false],
    ['2. Authorize', 'Tick B2 first.'],
    ['3. Install trigger', 'Complete B3 first.'],
    ['Before finishing', 'Read E2 and tick E3 →'],
    ['4. Finish setup', false],
    ['Status', 'Start with B2 in your own copy.']
  ]);
  sheet.getRange(1, 4, 11, 2).setValues([
    ['Setup details', 'Public access, connection and help'],
    ['Public-data disclosure', SHEET_DB_SETUP_.disclosure],
    ['I agree to this public setup', false],
    ['Last error', ''], ['Connection JSON', ''],
    ['Timeslot URL — copy', ''], ['Dashboard URL — copy', ''],
    ['Links expire', ''], ['Database identity', book.getId()],
    ['Instructions', 'Browser: keep the same Incognito/Private session and copy URLs. Mobile: edit checkboxes in Sheets using the copy-owner account. B2 → B3 → B4 → read E2 → tick E3 → wait for B7 → B6.'],
    ['Setup completed', '']
  ]);
  [2, 6].forEach(function (row) { sheet.getRange(row, 2).insertCheckboxes(); });
  sheet.getRange(3, 5).insertCheckboxes();
  PropertiesService.getScriptProperties().setProperty(SHEET_DB_SETUP_.property,
    JSON.stringify({ databaseId: book.getId(), sheetId: sheet.getSheetId() }));
  if (previousConfig && previousConfig.spreadsheetId === book.getId()) {
    sheetDbSetupPublish_(sheet, previousConfig);
    sheet.getRange(11, 5).setValue(completed);
    sheet.getRange(7, 2).setValue('Layout updated. Connection details preserved in E5–E7.');
  }
  sheetDbSetupFormatTemplate_(sheet);
  book.setActiveSheet(sheet);
}

function sheetDbSetupFormatTemplate_(sheet) {
  sheet.getRange(1, 1, 16, 5).setWrap(true);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e4efdf');
  sheet.getRange(1, 4, 1, 2).setFontWeight('bold').setBackground('#e4efdf');
  sheet.getRange(2, 4, 2, 2).setBackground('#fff0d4');
  sheet.setColumnWidth(1, 170).setColumnWidth(2, 300).setColumnWidth(3, 28)
    .setColumnWidth(4, 190).setColumnWidth(5, 370);
  sheet.setFrozenRows(1);
  sheet.setRowHeights(1, 16, 48);
  sheet.setRowHeight(2, 120).setRowHeight(5, 100).setRowHeight(10, 120);
  sheet.getRange(3, 2, 2, 1).setShowHyperlink(false).setWrap(false);
  sheet.getRange(5, 5, 3, 1).setWrap(false);
  // Checkbox glyphs use cell font sizing; keep labels and URL text compact.
  [[2, 2], [6, 2], [3, 5]].forEach(function (cell) {
    sheet.getRange(cell[0], cell[1]).setFontSize(32)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  sheet.setRowHeight(3, 64).setRowHeight(6, 64);
}

function sheetDbOpenSetup() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  sheetDbAssert_(book, 'SETUP_REQUIRED', 'Open the database spreadsheet first.');
  if (!book.getSheetByName(SHEET_DB_SETUP_.sheetName)) sheetDbPrepareSetupTemplate();
  book.setActiveSheet(sheetDbSetupSheet_(book));
}

function sheetDbSetupSheet_(book) {
  const sheet = book.getSheetByName(SHEET_DB_SETUP_.sheetName);
  sheetDbAssert_(sheet && sheet.getRange(1, 1).getValue() === SHEET_DB_SETUP_.marker,
    'SETUP_CHANGED', 'The managed Setup tab is missing or changed. Keep existing data and restore the prepared template.');
  return sheet;
}
function sheetDbSetupRead_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(SHEET_DB_SETUP_.property) || 'null'); }
  catch (_) { return null; }
}
function sheetDbSetupSave_(state) {
  PropertiesService.getScriptProperties().setProperty(SHEET_DB_SETUP_.property, JSON.stringify(state));
}
function sheetDbSetupOwns_(state, book, sheet) {
  return !!(state && state.databaseId === book.getId() && state.sheetId === sheet.getSheetId());
}
function sheetDbSetupCell_(event, row, column) {
  return !!(event && event.source && event.range && event.range.getSheet().getName() === SHEET_DB_SETUP_.sheetName &&
    event.range.getNumRows() === 1 && event.range.getNumColumns() === 1 &&
    event.range.getColumn() === (column || 2) && event.range.getRow() === row);
}
function sheetDbSetupUrl_(sheet, row, column, url) {
  sheet.getRange(row, column).setNumberFormat('@').setValue(url).setShowHyperlink(false);
}

function sheetDbSetupClearOutput_(sheet) {
  [5, 6, 7, 11].forEach(function (row) { sheet.getRange(row, 5).clearContent(); });
}
function sheetDbSetupError_(sheet, stage, error) {
  sheet.getRange(4, 5).setValue(stage + ': ' + String(error && error.message || error));
  sheet.getRange(7, 2).setValue('Setup stopped. Read E4 — Last error, correct it, then retry the same copy.');
}

/** Called by the sole simple onEdit in Code.gs. No authorized setup work here. */
function sheetDbSetupOnEdit_(event) {
  if (!event || !event.range || event.range.getSheet().getName() !== SHEET_DB_SETUP_.sheetName) return false;
  if (!sheetDbSetupCell_(event, 2) || event.value !== 'TRUE') return true;
  const sheet = event.range.getSheet(), lock = LockService.getScriptLock();
  let locked = false, stage = 'Prepare links';
  try {
    locked = lock.tryLock(5000);
    sheetDbAssert_(locked, 'BUSY', 'Setup is already running. Wait, then tick B2 again.');
    sheetDbSetupSheet_(event.source);
    sheet.getRange(2, 2).setValue(false);
    sheet.getRange(6, 2).setValue(false);
    const previous = sheetDbSetupRead_();
    const state = sheetDbSetupOwns_(previous, event.source, sheet) ? previous :
      { databaseId: event.source.getId(), sheetId: sheet.getSheetId() };
    if (state !== previous) {
      sheetDbSetupClearOutput_(sheet);
      sheet.getRange(3, 5).setValue(false);
    } else if (state.consentFor !== event.source.getId()) sheet.getRange(3, 5).setValue(false);
    sheet.getRange(9, 5).setValue(event.source.getId());
    sheet.getRange(2, 5).setValue(SHEET_DB_SETUP_.disclosure);
    [3, 4].forEach(function (row) { sheet.getRange(row, 2).clearContent(); });
    [4, 8].forEach(function (row) { sheet.getRange(row, 5).clearContent(); });
    sheet.getRange(7, 2).setValue('Copy B3 into your private browser to authorize, then copy B4 to install.');
    stage = 'Create authorization link';
    const authUrl = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationUrl();
    if (authUrl) sheetDbSetupUrl_(sheet, 3, 2, authUrl);
    else sheet.getRange(3, 2).setValue('Already authorized. Continue with B4.');
    stage = 'Create installation link';
    state.pending = { nonce: Utilities.getUuid(), expiresAt: Date.now() + SHEET_DB_SETUP_.lifetimeSeconds * 1000 };
    const token = ScriptApp.newStateToken().withMethod('sheetDbSetupInstall')
      .withArgument('nonce', state.pending.nonce).withTimeout(SHEET_DB_SETUP_.lifetimeSeconds).createToken();
    const url = 'https://script.google.com/macros/d/' + encodeURIComponent(ScriptApp.getScriptId()) +
      '/usercallback?state=' + encodeURIComponent(token);
    sheetDbSetupSave_(state);
    sheetDbSetupUrl_(sheet, 4, 2, url);
    sheet.getRange(8, 5).setValue(new Date(state.pending.expiresAt).toISOString());
  } catch (error) { sheetDbSetupError_(sheet, stage, error); }
  finally { if (locked) lock.releaseLock(); }
  return true;
}

function sheetDbSetupTriggers_(bookId) {
  return ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction() === SHEET_DB_SETUP_.handler && trigger.getTriggerSourceId() === bookId &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT;
  });
}
function sheetDbSetupPage_(title, message) {
  const escape = function (value) { return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  return HtmlService.createHtmlOutput('<!doctype html><main style="font:18px/1.5 system-ui;padding:24px">' +
    '<h1>' + escape(title) + '</h1><p>' + escape(message) + '</p></main>')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Signed browser callback: no active spreadsheet, openById, or spreadsheet UI. */
function sheetDbSetupInstall(event) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    locked = lock.tryLock(5000);
    sheetDbAssert_(locked, 'BUSY', 'Another installation is running. Wait, then reopen this link.');
    const state = sheetDbSetupRead_();
    sheetDbAssert_(state && state.pending && event && event.parameter && event.parameter.nonce === state.pending.nonce &&
      Number.isFinite(state.pending.expiresAt) && state.pending.expiresAt > Date.now() &&
      typeof state.databaseId === 'string' && /^[A-Za-z0-9_-]+$/.test(state.databaseId) && Number.isInteger(state.sheetId),
      'LINK_EXPIRED', 'The link expired, was replaced, or was used. Return to your copy and tick B2 for fresh links.');
    const auth = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    sheetDbAssert_(auth.getAuthorizationStatus() !== ScriptApp.AuthorizationStatus.REQUIRED,
      'AUTHORIZATION_REQUIRED', 'Complete B3 authorization for this copy before opening B4. Grant all requested permissions.');
    const matching = sheetDbSetupTriggers_(state.databaseId);
    if (!matching.length) ScriptApp.newTrigger(SHEET_DB_SETUP_.handler).forSpreadsheet(state.databaseId).onEdit().create();
    matching.slice(1).forEach(function (trigger) { ScriptApp.deleteTrigger(trigger); });
    delete state.pending;
    sheetDbSetupSave_(state);
    return sheetDbSetupPage_('Setup trigger installed',
      'Return to the Setup tab in the same copy. Read E2, tick E3 to agree, and wait for B7 to acknowledge it before ticking B6. ' +
      'Installation alone does not mean the database is ready. Wait for Setup complete and an app link in the Sheet.');
  } catch (error) {
    return sheetDbSetupPage_('Setup trigger not installed', String(error && error.message || error) +
      ' Open B3 first in your browser, signed in only as this copy’s owner, then reopen the current B4 link in that same session.');
  } finally { if (locked) lock.releaseLock(); }
}

function sheetDbSetupConsent_(book, sheet, triggerUid) {
  const state = sheetDbSetupRead_();
  return sheetDbSetupOwns_(state, book, sheet) && state.consentFor === book.getId() &&
    state.consentTriggerUid === String(triggerUid) && sheet.getRange(3, 5).getValue() === true &&
    sheet.getRange(2, 5).getValue() === SHEET_DB_SETUP_.disclosure;
}
function sheetDbSetupPublish_(sheet, config) {
  const json = JSON.stringify(config);
  sheet.getRange(5, 5).setNumberFormat('@').setValue(json);
  sheetDbSetupUrl_(sheet, 6, 5, SHEET_DB_SETUP_.apps.timeslot + '#sheetdb-forms=' + encodeURIComponent(json));
  sheetDbSetupUrl_(sheet, 7, 5, SHEET_DB_SETUP_.apps.dashboard + '#sheetdb-forms=' + encodeURIComponent(json));
  sheet.getRange(11, 5).setValue(new Date().toISOString());
  sheet.getRange(4, 5).clearContent();
  sheet.getRange(7, 2).setValue('Setup complete. Copy the app URL from E6 or E7 into your private browser.');
}

/** Installed edit trigger. Only its verified event may accept consent or initialize. */
function sheetDbSetupEdit(event) {
  const consent = sheetDbSetupCell_(event, 3, 5), finish = sheetDbSetupCell_(event, 6) && event.value === 'TRUE';
  if ((!consent && !finish) || !event.triggerUid) return;
  const sheet = event.range.getSheet(), book = event.source;
  let lock, locked = false;
  try {
    sheetDbSetupSheet_(book);
    const matching = sheetDbSetupTriggers_(book.getId());
    sheetDbAssert_(matching.some(function (trigger) { return String(trigger.getUniqueId()) === String(event.triggerUid); }),
      'TRIGGER_MISMATCH', 'This event does not match this account’s setup trigger. Start again at B2.');
    const state = sheetDbSetupRead_();
    sheetDbAssert_(sheetDbSetupOwns_(state, book, sheet), 'COPY_SETUP_REQUIRED', 'Tick B2 in this copy before authorizing or finishing setup.');
    if (consent) {
      lock = LockService.getScriptLock(); locked = lock.tryLock(5000);
      sheetDbAssert_(locked, 'BUSY', 'Setup is running. Wait, then change E3 again.');
      // Read again under the lock so link regeneration cannot be overwritten.
      const current = sheetDbSetupRead_();
      sheetDbAssert_(sheetDbSetupOwns_(current, book, sheet), 'COPY_SETUP_REQUIRED', 'Start this copy at B2.');
      if (event.value === 'TRUE' && sheet.getRange(3, 5).getValue() === true &&
          sheet.getRange(2, 5).getValue() === SHEET_DB_SETUP_.disclosure) {
        current.consentFor = book.getId(); current.consentTriggerUid = String(event.triggerUid);
        sheet.getRange(7, 2).setValue('Public setup acknowledged for this copy. Tick B6 to finish.');
      } else {
        delete current.consentFor; delete current.consentTriggerUid;
        sheet.getRange(7, 2).setValue('Public setup is not acknowledged. Read the disclosure and tick E3 before B6.');
      }
      sheetDbSetupSave_(current);
      sheet.getRange(4, 5).clearContent();
      return;
    }
    if (sheet.getRange(6, 2).getValue() !== true) return;
    sheet.getRange(6, 2).setValue(false);
    sheetDbSetupClearOutput_(sheet);
    const confirmed = function () { return sheetDbSetupConsent_(book, sheet, event.triggerUid); };
    sheetDbAssert_(confirmed(), 'PUBLIC_CONFIRMATION_REQUIRED',
      'After installing the trigger, read the disclosure and tick E3 in this copy before B6. If E3 was already checked, untick it and tick it again.');
    sheet.getRange(7, 2).setValue('Creating the Form and background processor. Keep this copy; retries reuse its resources.');
    const config = sheetDbSetupFormsPoc_(book, confirmed);
    sheetDbSetupPublish_(sheet, config);
  } catch (error) { sheetDbSetupError_(sheet, 'Finish setup', error); }
  finally { if (locked) lock.releaseLock(); }
}
