/**
 * Community event reimbursement: Google Apps Script backend
 * -----------------------------------------------------------
 * Receives claims from the LINE LIFF form, checks the LINE login,
 * saves receipts to Google Drive and records to this Google Sheet.
 *
 * HOW TO USE: Google Sheet > Extensions > Apps Script > paste this whole file.
 * Then follow the step-by-step guide (change LIFF_ID below, run setup once, deploy).
 */

// 1) PASTE YOUR LIFF ID HERE (looks like 1234567890-AbCdEfGh)
const LIFF_ID = 'PASTE-YOUR-LIFF-ID-HERE';

// 2) Optional settings
const TIMEZONE = 'Asia/Bangkok';
const MAX_RECEIPT_MB = 8;        // per file
const MAX_ITEMS = 10;            // records per claim
const CATEGORIES = ['venue', 'food', 'supplies', 'print', 'transport', 'other'];
const FOLDER_NAME = 'Reimbursement receipts';

const TAB = { events: 'Events', claims: 'Claims', expenses: 'Expenses' };
const HEAD = {
  events:   ['id', 'name', 'date', 'budget', 'status'],
  claims:   ['submitted_at', 'reference', 'status', 'event_id', 'event_name', 'payee_name', 'promptpay',
             'total_thb', 'items', 'note', 'line_user_id', 'line_name', 'receipts_folder', 'approver',
             'decision_date', 'submission_key'],
  expenses: ['reference', 'no', 'receipt_date', 'category', 'description', 'amount_thb', 'receipt_link']
};

/* ===================== ONE-TIME SETUP (run from the editor) ===================== */
function setup() {
  const ss = SpreadsheetApp.getActive();

  const events = ensureTab_(ss, TAB.events, HEAD.events);
  const claims = ensureTab_(ss, TAB.claims, HEAD.claims);
  ensureTab_(ss, TAB.expenses, HEAD.expenses);

  if (events.getLastRow() < 2) {   // one sample row to edit
    events.getRange(2, 1, 1, 5).setValues([['EV-2026-11', 'Community Meetup - November', '2026-11-21', 5000, 'Approved']]);
  }
  events.getRange('C2:C500').setNumberFormat('@');            // keep dates as plain text
  events.getRange('E2:E500').setDataValidation(list_(['Approved', 'Closed']));
  claims.getRange('C2:C5000').setDataValidation(list_(['Pending', 'Approved', 'Rejected', 'Paid']));
  claims.getRange('G2:G5000').setNumberFormat('@');          // keep leading 0 of PromptPay numbers
  claims.setFrozenRows(1);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FOLDER_ID')) {
    props.setProperty('FOLDER_ID', DriveApp.createFolder(FOLDER_NAME).getId());
  }
  ss.toast('Setup finished. Tabs created and Drive folder "' + FOLDER_NAME + '" is ready.', 'Reimbursement', 8);
}

function ensureTab_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function list_(values) {
  return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
}

/* ===================== WEB APP ===================== */
// The form asks for the list of approved events (shown in its drop-down)
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;
    if (action === 'events') return json_(readEvents_());
    return json_({ ok: true, message: 'Reimbursement service is running' });
  } catch (err) {
    return json_({ ok: false, error: message_(err) });
  }
}

// The form sends a claim
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const user = verifyLine_(body.accessToken);       // who is really sending this
    const clean = validate_(body);
    const out = saveClaim_(clean, user);
    return json_({ ok: true, reference: out.reference, duplicate: out.duplicate });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: message_(err) });
  }
}

/* ===================== LINE CHECK ===================== */
function verifyLine_(token) {
  if (!token) throw new Error('Please open this form from LINE.');
  const channelId = String(LIFF_ID).split('-')[0];
  const v = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true });
  if (v.getResponseCode() !== 200) throw new Error('LINE login could not be verified.');
  const info = JSON.parse(v.getContentText());
  if (String(info.client_id) !== channelId || Number(info.expires_in) <= 0) {
    throw new Error('LINE login does not belong to this app.');
  }
  const p = UrlFetchApp.fetch('https://api.line.me/v2/profile',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (p.getResponseCode() !== 200) throw new Error('Could not read the LINE profile.');
  const prof = JSON.parse(p.getContentText());
  return { userId: prof.userId, displayName: prof.displayName };
}

/* ===================== VALIDATION ===================== */
function readEvents_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.events);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(function (r) { return r[0] && String(r[4]).toLowerCase() === 'approved'; })
    .map(function (r) {
      const d = r[2] instanceof Date ? Utilities.formatDate(r[2], TIMEZONE, 'yyyy-MM-dd') : String(r[2] || '');
      return { id: String(r[0]), name: String(r[1]), date: d, budget: Number(r[3]) || 0 };
    });
}

function validate_(b) {
  const ev = readEvents_().filter(function (x) { return x.id === String((b.event || {}).id); })[0];
  if (!ev) throw new Error('This event is not open for claims.');

  const name = String((b.payee || {}).name || '').trim();
  if (name.length < 2 || name.length > 80) throw new Error('Payee name is missing.');
  const pp = String((b.payee || {}).promptpay || '').replace(/\D/g, '');
  if (!/^(0\d{9}|\d{13}|\d{15})$/.test(pp)) throw new Error('PromptPay number is not valid.');

  const items = b.items || [];
  const receipts = b.receipts || [];
  if (items.length < 1 || items.length > MAX_ITEMS) throw new Error('A claim needs 1 to ' + MAX_ITEMS + ' expenses.');

  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  let totalSatang = 0;
  const rows = items.map(function (it, i) {
    const amt = Math.round(Number(it.amount) * 100);
    if (!isFinite(amt) || amt <= 0 || amt > 100000000) throw new Error('Expense ' + (i + 1) + ': amount is not valid.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(it.date)) || it.date > today) throw new Error('Expense ' + (i + 1) + ': date is not valid.');
    if (CATEGORIES.indexOf(it.category) < 0) throw new Error('Expense ' + (i + 1) + ': category is missing.');
    const desc = String(it.description || '').trim();
    if (!desc) throw new Error('Expense ' + (i + 1) + ': description is missing.');
    const rc = receipts.filter(function (r) { return r.field === it.receiptField; })[0];
    if (!rc || !rc.data) throw new Error('Expense ' + (i + 1) + ': receipt is missing.');
    if (['image/jpeg', 'image/png', 'image/webp', 'application/pdf'].indexOf(rc.type) < 0) {
      throw new Error('Expense ' + (i + 1) + ': receipt must be a photo or PDF.');
    }
    if (rc.data.length * 0.75 > MAX_RECEIPT_MB * 1048576) throw new Error('Expense ' + (i + 1) + ': receipt is too large.');
    totalSatang += amt;
    return { no: i + 1, date: it.date, category: it.category, description: desc.slice(0, 200), amountSatang: amt, receipt: rc };
  });

  return {
    key: String(b.idempotencyKey || '').slice(0, 80),
    event: ev,
    payeeName: name,
    promptpay: pp,
    note: String(b.note || '').slice(0, 500),
    totalSatang: totalSatang,
    items: rows
  };
}

/* ===================== SAVE ===================== */
function saveClaim_(c, user) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let folder = null;
  try {
    const ss = SpreadsheetApp.getActive();
    const claims = ss.getSheetByName(TAB.claims);
    const expenses = ss.getSheetByName(TAB.expenses);
    if (!claims || !expenses) throw new Error('Setup has not been run yet.');

    // Same submission sent twice (e.g. bad network)? Return the first reference.
    if (c.key && claims.getLastRow() > 1) {
      const keys = claims.getRange(2, 16, claims.getLastRow() - 1, 1).getValues();
      const refs = claims.getRange(2, 2, claims.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === c.key) return { reference: refs[i][0], duplicate: true };
      }
    }

    const props = PropertiesService.getScriptProperties();
    const seq = Number(props.getProperty('SEQ') || '0') + 1;
    const now = new Date();
    const reference = 'RC-' + Utilities.formatDate(now, TIMEZONE, 'yyMM') + '-' + ('0000' + seq).slice(-4);

    const root = DriveApp.getFolderById(props.getProperty('FOLDER_ID'));
    folder = root.createFolder((reference + ' ' + c.payeeName).replace(/[\\/:*?"<>|]/g, ' '));

    const expRows = c.items.map(function (it) {
      const ext = it.receipt.type === 'application/pdf' ? 'pdf' : (it.receipt.type === 'image/png' ? 'png' : 'jpg');
      const blob = Utilities.newBlob(Utilities.base64Decode(it.receipt.data), it.receipt.type,
        reference + '-' + it.no + '-' + it.category + '.' + ext);
      const file = folder.createFile(blob);
      return [reference, it.no, it.date, it.category, it.description, it.amountSatang / 100, file.getUrl()];
    });
    expenses.getRange(expenses.getLastRow() + 1, 1, expRows.length, 7).setValues(expRows);

    const row = [now, reference, 'Pending', c.event.id, c.event.name, c.payeeName, c.promptpay,
      c.totalSatang / 100, c.items.length, c.note, user.userId, user.displayName, folder.getUrl(), '', '', c.key];
    const r = claims.getLastRow() + 1;
    claims.getRange(r, 7).setNumberFormat('@');
    claims.getRange(r, 1, 1, row.length).setValues([row]);
    claims.getRange(r, 1).setNumberFormat('yyyy-mm-dd hh:mm');

    props.setProperty('SEQ', String(seq));
    return { reference: reference, duplicate: false };
  } catch (err) {
    if (folder) { try { folder.setTrashed(true); } catch (e) { /* ignore */ } }   // do not leave half-saved receipts
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/* ===================== HELPERS ===================== */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function message_(err) { return String((err && err.message) || err); }
