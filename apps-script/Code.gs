/**
 * Pat & Pete RSVP endpoint (v3: guest confirmation e-mail with e-cards).
 * Paste into the Google Sheet: Extensions > Apps Script, then Deploy > Manage deployments > pencil > New version > Deploy.
 */
const SHEET_NAME = 'Responses';
const NOTIFY_EMAIL = 'peerakarn@ionenergy.ai';        // your own alert
const FROM_NAME = 'Pat & Pete';                        // sender name guests see
const REPLY_TO = 'peerakarn@ionenergy.ai';
const SITE_URL = 'https://petemanakit.github.io/wedding/';
const ASSET_BASE = 'https://petemanakit.github.io/wedding/assets/';
const HEADERS = ['Timestamp', 'Invitation', 'Household / name', 'Guest names', 'Email', 'Phone',
  'Rehearsal dinner (Sat 16 Jan)', 'Vow ceremony (Sun 4.30 PM)', 'Evening celebration (Sun 6.30 PM)',
  'Headcount', 'Dietary requirements', 'Message', 'Kind', 'Drinks', 'Confirmation sent'];

const EVENTS = [
  { key: 'rehearsal', tier: 3, label: 'Rehearsal dinner', when: 'Saturday 16 January, 5.30 PM', where: 'Heritage Ballroom', card: 'ecard-rehearsal.jpg', cardName: 'Rehearsal Dinner.jpg' },
  { key: 'vow',       tier: 2, label: 'Vow ceremony',     when: 'Sunday 17 January, 4.30 PM',   where: 'Riverside Lawn',    card: 'ecard-vow.jpg',       cardName: 'Vow Ceremony.jpg' },
  { key: 'evening',   tier: 1, label: 'Wedding celebration', when: 'Sunday 17 January, 6.30 PM', where: 'Heritage Ballroom', card: 'ecard-evening.jpg',  cardName: 'Wedding Celebration.jpg' },
];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME, 0); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#595D44').setFontColor('#F7F4EE');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, HEADERS.length, 160); sh.setColumnWidth(12, 320);
    ensureSummary_(ss);
  }
  // add any headers introduced by a later version
  const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  HEADERS.forEach(function (h, i) {
    if (have[i] !== h && have.indexOf(h) === -1) {
      sh.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#595D44').setFontColor('#F7F4EE');
    }
  });
  return sh;
}

function ensureSummary_(ss) {
  if (ss.getSheetByName('Summary')) return;
  const s = ss.insertSheet('Summary');
  const r = "'" + SHEET_NAME + "'!";
  s.getRange('A1:C1').setValues([['Event', 'Accepting (households)', 'Guests (headcount)']]).setFontWeight('bold');
  s.getRange('A2:A4').setValues([['Rehearsal dinner'], ['Vow ceremony'], ['Evening celebration']]);
  s.getRange('B2:C4').setFormulas([
    ['=COUNTIF(' + r + 'G:G,"Accept")', '=SUMIF(' + r + 'G:G,"Accept",' + r + 'J:J)'],
    ['=COUNTIF(' + r + 'H:H,"Accept")', '=SUMIF(' + r + 'H:H,"Accept",' + r + 'J:J)'],
    ['=COUNTIF(' + r + 'I:I,"Accept")', '=SUMIF(' + r + 'I:I,"Accept",' + r + 'J:J)'],
  ]);
  s.getRange('A6').setValue('Total replies'); s.getRange('B6').setFormula('=COUNTA(' + r + 'A:A)-1');
  s.getRange('A7').setValue('Declined all');   s.getRange('B7').setFormula('=COUNTIFS(' + r + 'I:I,"Decline")');
  s.autoResizeColumns(1, 3);
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return json_({ ok: true }); // honeypot filled: silently ignore
    let confirmation = '';
    try { confirmation = confirmGuest_(p) ? 'Sent' : 'No e-mail'; } catch (err) { confirmation = 'Failed: ' + err; }
    const row = [
      new Date(), p.tier || '', p.household || '', p.guests || '', p.email || '', p.phone || '',
      p.rehearsal || '', p.vow || '', p.evening || '',
      Number(p.headcount || 0), p.dietary || '', p.message || '', p.kind || 'New', p.drinks || '', confirmation
    ];
    const sh = sheet_(); const r = sh.getLastRow() + 1;
    sh.getRange(r, 2, 1, HEADERS.length - 1).setNumberFormat('@'); // text, so '+66…' is never read as a formula
    sh.getRange(r, 10).setNumberFormat('0');
    sh.getRange(r, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    notify_(p, confirmation);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return ContentService.createTextOutput('Pat & Pete RSVP endpoint is live (v3).'); }

/* ---------- your own alert ---------- */
function notify_(p, confirmation) {
  if (!NOTIFY_EMAIL) return;
  const lines = [
    'Household: ' + (p.household || ''), 'Guests: ' + (p.guests || ''), 'Headcount: ' + (p.headcount || ''),
    'Invitation: tier ' + (p.tier || ''), 'Rehearsal dinner: ' + (p.rehearsal || 'n/a'),
    'Vow ceremony: ' + (p.vow || 'n/a'), 'Evening: ' + (p.evening || ''),
    'Email: ' + (p.email || ''), 'Phone: ' + (p.phone || ''), 'Dietary: ' + (p.dietary || ''),
    'Drinks: ' + (p.drinks || ''), 'Message: ' + (p.message || ''),
    'Confirmation to guest: ' + confirmation, '', 'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ];
  const verdict = accepted_(p) ? 'accepts' : 'declines';
  MailApp.sendEmail(NOTIFY_EMAIL, 'RSVP: ' + (p.household || 'Guest') + ' ' + verdict, lines.join('\n'));
}

/* ---------- guest confirmation ---------- */
function accepted_(p) { return p.evening === 'Accept' || p.vow === 'Accept' || p.rehearsal === 'Accept'; }

function confirmGuest_(p) {
  const to = String(p.email || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) return false;
  const tier = Number(p.tier || 1);
  const invited = EVENTS.filter(function (ev) { return ev.tier <= tier; });
  const names = (p.guests || '').trim() || 'Guests';
  const first = (p.household || names).trim();
  const yes = accepted_(p);

  const subject = yes
    ? 'Thank you, ' + first + ': your reply for 17 January is received'
    : 'Thank you, ' + first + ': we will miss you in January';
  const intro = yes
    ? 'Thank you for your reply. We are delighted that you will be with us in January, and we have noted the following.'
    : 'Thank you for letting us know. We shall miss you in January and hope to raise a glass with you before long. Your reply is noted below.';

  const rows = invited.map(function (ev) {
    const v = p[ev.key] === 'Accept' ? 'Joyfully accepted' : (p[ev.key] === 'Decline' ? 'Regretfully declined' : 'No reply given');
    return tr_(ev.label + ' &middot; ' + esc_(ev.when), esc_(v), '#595D44');
  }).join('');
  const details =
    tr_('Guests', esc_(names) + (p.headcount ? ' (' + esc_(p.headcount) + ')' : '')) +
    tr_('Dietary requirements', esc_(p.dietary || 'None')) +
    tr_('Preferred drinks', esc_(p.drinks || 'No preference given'));

  const html =
    '<div style="background:#F7F4EE;padding:28px 16px;font-family:Georgia,\'Times New Roman\',serif;color:#33352A">' +
    '<div style="max-width:560px;margin:0 auto;background:#FFFDF9;border:1px solid #D9D1BF">' +
    '<div style="background:#595D44;padding:28px 24px;text-align:center">' +
    '<img src="' + ASSET_BASE + 'monogram.png" width="64" alt="" style="width:64px;height:auto">' +
    '<div style="font-family:Georgia,serif;font-style:italic;color:#F7F4EE;font-size:30px;line-height:1.1;margin-top:10px">Pat &amp; Pete</div>' +
    '<div style="color:#E4D3A6;font-size:10px;letter-spacing:.3em;margin-top:8px">THE LANGHAM, CUSTOM HOUSE, BANGKOK &middot; 17 . 01 . 2027</div></div>' +
    '<div style="padding:30px 30px 10px;font-size:17px;line-height:1.55">' +
    '<p style="margin:0 0 14px">Dear ' + esc_(names) + ',</p>' +
    '<p style="margin:0 0 14px">' + intro + '</p>' +
    label_('Your reply') + '<table style="width:100%;border-collapse:collapse">' + rows + '</table>' +
    label_('Details') + '<table style="width:100%;border-collapse:collapse">' + details + '</table>' +
    '<p style="margin:22px 0 14px">Should anything change, simply reply to this e-mail or send a fresh reply through the website; we will use the latest one. All the details of the weekend, including the dress code and where to stay, remain on the website' +
    (invited.length ? ', and your invitation card' + (invited.length > 1 ? 's are' : ' is') + ' attached for your keeping.' : '.') + '</p>' +
    '</div>' +
    '<div style="text-align:center;padding:10px 30px 30px"><a href="' + SITE_URL + '" style="display:inline-block;background:#595D44;color:#F7F4EE;font-size:11px;letter-spacing:.26em;text-transform:uppercase;padding:14px 28px;text-decoration:none">Visit the wedding website</a></div>' +
    '<div style="border-top:1px solid #EFEAE0;padding:18px 30px 24px;text-align:center;font-size:14px;color:#6F7160">With love, Pat &amp; Pete<br>Questions? Write to <a href="mailto:' + REPLY_TO + '" style="color:#B08D3E">' + REPLY_TO + '</a></div>' +
    '</div></div>';

  const text = 'Dear ' + names + ',\n\n' + intro + '\n\n' +
    invited.map(function (ev) { return ev.label + ' (' + ev.when + '): ' + (p[ev.key] || 'no reply given'); }).join('\n') +
    '\n\nGuests: ' + names + (p.headcount ? ' (' + p.headcount + ')' : '') + '\nDietary: ' + (p.dietary || 'None') + '\nDrinks: ' + (p.drinks || 'No preference given') +
    '\n\nWebsite: ' + SITE_URL + '\n\nWith love, Pat & Pete';

  // attach only the e-cards for the events this guest is invited to
  const attachments = [];
  invited.forEach(function (ev) {
    try {
      const res = UrlFetchApp.fetch(ASSET_BASE + ev.card, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) attachments.push(res.getBlob().setName(ev.cardName));
    } catch (err) { /* skip a card that cannot be fetched; the e-mail still goes */ }
  });

  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, body: text, name: FROM_NAME, replyTo: REPLY_TO, attachments: attachments });
  return true;
}

function tr_(k, v, color) {
  return '<tr><td style="padding:9px 0;border-bottom:1px solid #EFEAE0;color:#6F7160;font-size:16px;width:46%;vertical-align:top">' + k + '</td>' +
    '<td style="padding:9px 0;border-bottom:1px solid #EFEAE0;font-weight:bold;color:' + (color || '#33352A') + ';font-size:16px;vertical-align:top">' + v + '</td></tr>';
}
function label_(t) { return '<div style="font-size:10px;letter-spacing:.28em;text-transform:uppercase;color:#B08D3E;margin:24px 0 8px">' + t + '</div>'; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
