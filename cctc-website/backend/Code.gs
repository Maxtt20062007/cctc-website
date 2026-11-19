/**
 * CCTC Lead System — Google Apps Script
 * --------------------------------------------------------------
 * What it does, automatically:
 *  1. Receives every website form request and saves it to the "Leads" sheet.
 *  2. Emails YOU instantly with the customer's details (tap-to-call link).
 *  3. Emails the CUSTOMER a confirmation (if they gave an email).
 *  4. Sets a "Call by" date: today if the request comes before the cutoff
 *     hour on a work day, otherwise the next work day.
 *  5. Every work-day morning, emails you a to-do list: who to call,
 *     who to call back, survey visits, quote follow-ups, review requests,
 *     plus your daily checklist.
 *  6. When you change a lead's Status, it sets the next follow-up date for you.
 *  7. Adds a "CCTC" menu to the sheet to send follow-up emails you approve.
 *
 * Setup: see README.md. Edit SETTINGS below, then run setup() once.
 */

const SETTINGS = {
  OWNER_EMAIL: 'YOUR_EMAIL@gmail.com',     // where new-lead alerts and the daily list go
  OWNER_NAME: 'Minh',
  BUSINESS_NAME: 'CCTC (Canadian Trusted Construction Company)',
  BUSINESS_PHONE: '',                      // e.g. '306-555-0123' (shown in customer emails)
  TIMEZONE: 'America/Regina',              // Saskatchewan time (no daylight saving)
  WORK_DAYS: [1, 2, 3, 4, 5, 6],           // 1=Mon ... 7=Sun. Mon–Sat by default.
  SAME_DAY_CUTOFF_HOUR: 15,                // requests before 3 pm on a work day → call same day
  DIGEST_HOUR: 7,                          // daily to-do email at about 7 am
  NO_ANSWER_RETRY_DAYS: 1,                 // "No answer" → try again next work day
  QUOTE_FOLLOWUP_DAYS: 3,                  // "Quoted" → follow up after 3 work days
  REVIEW_REQUEST_DAYS: 2,                  // "Done" → ask for a review after 2 work days
  SEND_CUSTOMER_CONFIRMATION: true,

  // Your daily routine. Edit freely; it appears at the bottom of each morning email.
  DAILY_CHECKLIST: [
    'Call every lead on today\'s list (log the result in the Status column)',
    'Confirm today\'s site surveys by text or phone',
    'Log yesterday\'s job: costs, revenue, margin, lessons (Step 05 · Record)',
    'Move 50% of any new profit to the reserve fund',
    'Write today\'s job journal entry',
  ],
  // Training focus by weekday (from your 90-day plan). Remove any you don't want.
  TRAINING: { 1: 'Marketing', 2: 'Websites & AI', 3: 'Accounting', 4: 'Time & reporting',
              5: 'Emotional discipline', 6: 'Weekly report', 7: '' },
  TRAINING_DAILY: 'Mandarin practice',
};

const SHEET_NAME = 'Leads';
const HEADERS = ['ID', 'Received', 'Name', 'Phone', 'Email', 'Address / area', 'Postal code',
  'Service', 'Best time to call', 'When needed', 'Job details', 'Status', 'Call by',
  'Next follow-up', 'Notes', 'Last updated'];
const COL = {}; HEADERS.forEach((h, i) => COL[h] = i + 1);
const STATUSES = ['New', 'No answer', 'Survey booked', 'Quoted', 'Won', 'Done', 'Lost'];

/* ============================ SETUP ============================ */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(SETTINGS.TIMEZONE);
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#0C1526').setFontColor('#FFB23F');
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL['Job details'], 320);
  sh.setColumnWidth(COL['Notes'], 260);

  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  sh.getRange(2, COL['Status'], sh.getMaxRows() - 1, 1).setDataValidation(statusRule);
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  [COL['Call by'], COL['Next follow-up']].forEach(c => {
    sh.getRange(2, c, sh.getMaxRows() - 1, 1).setDataValidation(dateRule).setNumberFormat('ddd mmm d, yyyy');
  });
  sh.getRange(2, COL['Received'], sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL['Last updated'], sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL['Phone'], sh.getMaxRows() - 1, 1).setNumberFormat('@');

  // Colour rows by status
  const rng = sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length);
  const L = columnLetter(COL['Status']);
  const colours = { 'New': '#FFF4DC', 'No answer': '#FFE3DF', 'Survey booked': '#E1F4FF', 'Quoted': '#EAE6FF', 'Won': '#DFF7EA', 'Done': '#EEEEEE', 'Lost': '#F5F5F5' };
  sh.setConditionalFormatRules(Object.keys(colours).map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + L + '2="' + s + '"')
      .setBackground(colours[s]).setRanges([rng]).build()));

  // Triggers: daily digest + status-change handler (remove old copies first)
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(SETTINGS.DIGEST_HOUR)
    .inTimezone(SETTINGS.TIMEZONE).create();
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();

  try {
    SpreadsheetApp.getUi().alert('CCTC setup complete.\n\nNext: Deploy → New deployment → Web app (see README step 4).');
  } catch (e) { /* run from editor without UI */ }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CCTC')
    .addItem('Email selected customer: quote follow-up', 'menuQuoteFollowUp')
    .addItem('Email selected customer: thank-you + review request', 'menuReviewRequest')
    .addSeparator()
    .addItem('Send me today\'s to-do email now', 'dailyDigestForce')
    .addItem('Add a test lead', 'testLead')
    .addItem('Run setup again (fix triggers/format)', 'setup')
    .addToUi();
}

/* ======================== WEBSITE FORM ========================= */

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json({ ok: true });                 // spam bot filled the hidden field
    const name = clean(d.name, 100), phone = clean(d.phone, 40), email = clean(d.email, 120);
    const details = clean(d.details, 2000);
    if (!name || !phone || !details) return json({ ok: false, error: 'missing fields' });
    if (phone.replace(/\D/g, '').length < 10) return json({ ok: false, error: 'bad phone' });

    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    const sh = leadsSheet();
    const id = 'L' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    const callBy = callByDate(new Date());
    const row = new Array(HEADERS.length).fill('');
    row[COL['ID'] - 1] = id;
    row[COL['Received'] - 1] = new Date();
    row[COL['Name'] - 1] = safe(name);
    row[COL['Phone'] - 1] = safe(phone);
    row[COL['Email'] - 1] = safe(email);
    row[COL['Address / area'] - 1] = safe(clean(d.address, 200));
    row[COL['Postal code'] - 1] = safe(clean(d.postal, 10).toUpperCase());
    row[COL['Service'] - 1] = safe(clean(d.service, 80));
    row[COL['Best time to call'] - 1] = safe(clean(d.bestTime, 40));
    row[COL['When needed'] - 1] = safe(clean(d.when, 40));
    row[COL['Job details'] - 1] = safe(details);
    row[COL['Status'] - 1] = 'New';
    row[COL['Call by'] - 1] = callBy;
    row[COL['Last updated'] - 1] = new Date();
    sh.appendRow(row);
    lock.releaseLock();

    const lead = rowToLead(row);
    notifyOwner(lead);
    if (SETTINGS.SEND_CUSTOMER_CONFIRMATION && validEmail(email)) confirmCustomer(lead);

    return json({ ok: true, id: id, callByText: friendlyDay(callBy) });
  } catch (err) {
    console.error(err);
    try { MailApp.sendEmail(SETTINGS.OWNER_EMAIL, 'CCTC website form error', String(err && err.stack || err)); } catch (_) {}
    return json({ ok: false, error: 'server error' });
  }
}

function doGet() { return ContentService.createTextOutput('CCTC lead endpoint is running.'); }

/* ========================== EMAILS ============================= */

function notifyOwner(l) {
  const tel = l.phone.replace(/[^\d+]/g, '');
  const html =
    '<h2 style="margin:0">New request: ' + esc(l.service) + '</h2>' +
    '<p><b>' + esc(l.name) + '</b> · <a href="tel:' + tel + '">' + esc(l.phone) + '</a>' +
    (l.email ? ' · <a href="mailto:' + esc(l.email) + '">' + esc(l.email) + '</a>' : '') + '</p>' +
    '<p><b>Call by:</b> ' + friendlyDay(l.callBy) + ' &nbsp; <b>Best time:</b> ' + esc(l.bestTime) + ' &nbsp; <b>When:</b> ' + esc(l.when) + '</p>' +
    '<p><b>Where:</b> ' + esc(l.address || '-') + ' ' + esc(l.postal || '') + '</p>' +
    '<p style="white-space:pre-wrap;border-left:3px solid #FFB23F;padding-left:10px">' + esc(l.details) + '</p>' +
    '<p><a href="' + sheetUrl() + '">Open the Leads sheet</a> · ID ' + l.id + '</p>';
  MailApp.sendEmail({ to: SETTINGS.OWNER_EMAIL, subject: '🔔 New lead: ' + l.name + ' – ' + l.service + ' (call by ' + friendlyDay(l.callBy) + ')', htmlBody: html, replyTo: l.email || SETTINGS.OWNER_EMAIL });
}

function confirmCustomer(l) {
  const first = l.name.split(/\s+/)[0];
  const body =
    'Hi ' + first + ',\n\n' +
    'Thanks for contacting ' + SETTINGS.BUSINESS_NAME + '. We received your request for: ' + l.service + '.\n\n' +
    'What happens next:\n' +
    '1. ' + SETTINGS.OWNER_NAME + ' will call you by ' + friendlyDay(l.callBy) + (l.bestTime && l.bestTime !== 'Anytime' ? ' (' + l.bestTime.toLowerCase() + ', as you asked)' : '') + '.\n' +
    '2. We book a site survey to see the job in person.\n' +
    '3. You get a written, line-by-line quote before any work starts.\n\n' +
    'Your request:\n"' + l.details + '"\n\n' +
    (SETTINGS.BUSINESS_PHONE ? 'Questions before then? Call ' + SETTINGS.BUSINESS_PHONE + '.\n\n' : 'Questions before then? Just reply to this email.\n\n') +
    SETTINGS.OWNER_NAME + '\n' + SETTINGS.BUSINESS_NAME + '\nMoose Jaw, SK\n\n(Reference: ' + l.id + ')';
  MailApp.sendEmail({ to: l.email, subject: 'We got your request – ' + SETTINGS.BUSINESS_NAME, body: body, replyTo: SETTINGS.OWNER_EMAIL, name: 'CCTC' });
}

/* ======================= DAILY TO-DO EMAIL ===================== */

function dailyDigest() { sendDigest(false); }
function dailyDigestForce() { sendDigest(true); try { SpreadsheetApp.getUi().alert('Sent to ' + SETTINGS.OWNER_EMAIL); } catch (_) {} }

function sendDigest(force) {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, SETTINGS.TIMEZONE, 'u'));
  if (!force && SETTINGS.WORK_DAYS.indexOf(dow) === -1) return;
  const today = dayKey(now);
  const leads = allLeads();

  const callNew = [], retry = [], surveys = [], quotes = [], reviews = [], unscheduled = [];
  let active = 0;
  leads.forEach(l => {
    const cb = l.callBy ? dayKey(l.callBy) : '', fu = l.nextFollowUp ? dayKey(l.nextFollowUp) : '';
    switch (l.status) {
      case 'New': if (!cb || cb <= today) callNew.push(l); break;
      case 'No answer': if (!fu || fu <= today) retry.push(l); break;
      case 'Survey booked': if (!fu) unscheduled.push(l); else if (fu <= today) surveys.push(l); break;
      case 'Quoted': if (fu && fu <= today) quotes.push(l); break;
      case 'Done': if (fu && fu <= today) reviews.push(l); break;
      case 'Won': active++; break;
    }
  });

  const total = callNew.length + retry.length + surveys.length + quotes.length + reviews.length;
  const item = (l, extra) => {
    const tel = l.phone.replace(/[^\d+]/g, '');
    const cb = l.callBy ? dayKey(l.callBy) : '';
    const late = l.status === 'New' && cb && cb < today ? ' <b style="color:#c0392b">OVERDUE</b>' : '';
    return '<li style="margin-bottom:8px"><b>' + esc(l.name) + '</b> – <a href="tel:' + tel + '">' + esc(l.phone) + '</a>' + late +
      '<br><span style="color:#555">' + esc(l.service) + (l.bestTime ? ' · best: ' + esc(l.bestTime) : '') +
      (l.address ? ' · ' + esc(l.address) : '') + (extra ? ' · ' + extra : '') + '</span>' +
      (l.notes ? '<br><i style="color:#777">Notes: ' + esc(l.notes) + '</i>' : '') + '</li>';
  };
  const section = (title, list, hint, extraFn) => list.length
    ? '<h3 style="margin:18px 0 6px">' + title + ' (' + list.length + ')</h3>' + (hint ? '<p style="margin:0 0 6px;color:#666">' + hint + '</p>' : '') +
      '<ul style="padding-left:18px;margin:0">' + list.map(l => item(l, extraFn ? extraFn(l) : '')).join('') + '</ul>'
    : '';

  const training = SETTINGS.TRAINING[dow];
  const checklist = SETTINGS.DAILY_CHECKLIST.slice();
  if (training) checklist.push('Training: ' + training);
  if (SETTINGS.TRAINING_DAILY) checklist.push(SETTINGS.TRAINING_DAILY);

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:620px">' +
    '<h2 style="margin:0">CCTC · ' + Utilities.formatDate(now, SETTINGS.TIMEZONE, 'EEEE, MMM d') + '</h2>' +
    '<p style="color:#555">' + (total ? total + ' thing' + (total > 1 ? 's' : '') + ' to do with customers today.' : 'No customer calls due today. 🎉') +
    (active ? ' Jobs in progress: ' + active + '.' : '') + '</p>' +
    section('📞 Call new leads', callNew, 'After each call, change Status to Survey booked, No answer, Quoted or Lost.') +
    section('🔁 Call back (no answer last time)', retry) +
    section('📐 Site surveys today', surveys) +
    section('💬 Quote follow-ups', quotes, 'Use the CCTC menu → "quote follow-up" to email them, or call.') +
    section('⭐ Ask for a review', reviews, 'Use the CCTC menu → "thank-you + review request".') +
    section('⚠️ Survey booked but no date set', unscheduled, 'Put the visit date in "Next follow-up".') +
    '<h3 style="margin:22px 0 6px">✅ Daily checklist</h3><ul style="padding-left:18px;margin:0">' +
    checklist.map(c => '<li>' + esc(c) + '</li>').join('') + '</ul>' +
    '<p style="margin-top:22px"><a href="' + sheetUrl() + '">Open the Leads sheet</a></p></div>';

  MailApp.sendEmail({ to: SETTINGS.OWNER_EMAIL, subject: 'CCTC today: ' + (callNew.length + retry.length) + ' calls, ' + (surveys.length) + ' surveys, ' + (quotes.length + reviews.length) + ' follow-ups', htmlBody: html });
}

/* =================== STATUS CHANGES (auto dates) ================ */

function handleEdit(e) {
  const r = e && e.range; if (!r) return;
  const sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME || r.getRow() < 2 || r.getColumn() !== COL['Status'] || r.getNumColumns() > 1) return;
  const row = r.getRow(), status = String(r.getValue());
  const fuCell = sh.getRange(row, COL['Next follow-up']);
  const today = new Date();
  switch (status) {
    case 'No answer': fuCell.setValue(addWorkDays(today, SETTINGS.NO_ANSWER_RETRY_DAYS)); break;
    case 'Quoted': fuCell.setValue(addWorkDays(today, SETTINGS.QUOTE_FOLLOWUP_DAYS)); break;
    case 'Done': fuCell.setValue(addWorkDays(today, SETTINGS.REVIEW_REQUEST_DAYS)); break;
    case 'Won': case 'Lost': fuCell.clearContent(); break;
    case 'Survey booked':
      if (!fuCell.getValue()) fuCell.setNote('Type the survey visit date here.');
      break;
  }
  sh.getRange(row, COL['Last updated']).setValue(new Date());
}

/* ============== MENU: follow-up emails you approve ============== */

function menuQuoteFollowUp() {
  sendTemplateToSelected('Following up on your quote – CCTC', l =>
    'Hi ' + l.name.split(/\s+/)[0] + ',\n\nI wanted to follow up on the quote for your ' + l.service.toLowerCase() +
    ' job. Do you have any questions, or would you like to pick a start date?\n\nHappy to adjust the scope if needed.\n\n' +
    SETTINGS.OWNER_NAME + '\n' + SETTINGS.BUSINESS_NAME + (SETTINGS.BUSINESS_PHONE ? '\n' + SETTINGS.BUSINESS_PHONE : ''),
    function (sh, row) { sh.getRange(row, COL['Next follow-up']).setValue(addWorkDays(new Date(), SETTINGS.QUOTE_FOLLOWUP_DAYS)); });
}

function menuReviewRequest() {
  sendTemplateToSelected('Thank you from CCTC', l =>
    'Hi ' + l.name.split(/\s+/)[0] + ',\n\nThank you for choosing CCTC for your ' + l.service.toLowerCase() +
    '. If anything about the work isn\'t right, reply and I\'ll come back to fix it.\n\n' +
    'If you were happy, a short review or a referral to a friend would mean a lot to a new local business.\n\n' +
    SETTINGS.OWNER_NAME + '\n' + SETTINGS.BUSINESS_NAME,
    function (sh, row) { sh.getRange(row, COL['Next follow-up']).clearContent(); });
}

function sendTemplateToSelected(subject, bodyFn, after) {
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_NAME) return ui.alert('Click a row in the Leads sheet first.');
  const row = sh.getActiveRange().getRow();
  if (row < 2) return ui.alert('Click a customer row first.');
  const l = rowToLead(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0]);
  if (!validEmail(l.email)) return ui.alert(l.name + ' has no email address. Give them a call instead: ' + l.phone);
  const body = bodyFn(l);
  const ok = ui.alert('Send this email to ' + l.name + ' <' + l.email + '>?', body, ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  MailApp.sendEmail({ to: l.email, subject: subject, body: body, replyTo: SETTINGS.OWNER_EMAIL, name: 'CCTC' });
  const notes = sh.getRange(row, COL['Notes']);
  notes.setValue((notes.getValue() ? notes.getValue() + '\n' : '') + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'MMM d') + ': emailed "' + subject + '"');
  if (after) after(sh, row);
  sh.getRange(row, COL['Last updated']).setValue(new Date());
  ui.alert('Sent.');
}

/* ============================ TEST ============================= */

function testLead() {
  const res = doPost({ postData: { contents: JSON.stringify({
    name: 'Test Customer', phone: '306-555-0100', email: SETTINGS.OWNER_EMAIL, address: '123 Main St N',
    postal: 'S6H 1A1', service: 'Interior painting', bestTime: 'Afternoon (12–5)', when: 'Within 2 weeks',
    details: 'TEST – repaint a 12 x 14 ft bedroom. Delete this row after testing.' }) } });
  console.log(res.getContent());
}

/* =========================== HELPERS =========================== */

function leadsSheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME);
}
function sheetUrl() { const sh = leadsSheet(); return sh.getParent().getUrl() + '#gid=' + sh.getSheetId(); }
function allLeads() {
  const sh = leadsSheet(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEADERS.length).getValues().filter(r => r[COL['Name'] - 1]).map(rowToLead);
}
function rowToLead(r) {
  const g = h => r[COL[h] - 1];
  return { id: String(g('ID')), name: String(g('Name')), phone: String(g('Phone')), email: String(g('Email')),
    address: String(g('Address / area')), postal: String(g('Postal code')), service: String(g('Service') || 'General request'),
    bestTime: String(g('Best time to call')), when: String(g('When needed')), details: String(g('Job details')),
    status: String(g('Status') || 'New'), callBy: g('Call by'), nextFollowUp: g('Next follow-up'), notes: String(g('Notes')) };
}
function isWorkDay(d) { const u = Number(Utilities.formatDate(d, SETTINGS.TIMEZONE, 'u')); return SETTINGS.WORK_DAYS.indexOf(u) !== -1; }
function startOfDay(d) { const s = Utilities.formatDate(d, SETTINGS.TIMEZONE, 'yyyy-MM-dd').split('-'); return new Date(+s[0], +s[1] - 1, +s[2], 12); }
function addWorkDays(d, n) { let x = startOfDay(d), k = 0; while (k < n) { x = new Date(x.getTime() + 864e5); if (isWorkDay(x)) k++; } return x; }
function callByDate(now) {
  const hour = Number(Utilities.formatDate(now, SETTINGS.TIMEZONE, 'H'));
  if (isWorkDay(now) && hour < SETTINGS.SAME_DAY_CUTOFF_HOUR) return startOfDay(now);
  return addWorkDays(now, 1);
}
function dayKey(d) { return d instanceof Date ? Utilities.formatDate(d, SETTINGS.TIMEZONE, 'yyyy-MM-dd') : String(d); }
function friendlyDay(d) {
  if (!(d instanceof Date)) return String(d || '');
  const k = dayKey(d), today = dayKey(new Date()), tmr = dayKey(new Date(Date.now() + 864e5));
  if (k === today) return 'today (' + Utilities.formatDate(d, SETTINGS.TIMEZONE, 'EEE MMM d') + ')';
  if (k === tmr) return 'tomorrow (' + Utilities.formatDate(d, SETTINGS.TIMEZONE, 'EEE MMM d') + ')';
  return Utilities.formatDate(d, SETTINGS.TIMEZONE, 'EEEE, MMM d');
}
function clean(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 500); }
function safe(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }            // stops spreadsheet formula injection
function validEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '')); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function columnLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
