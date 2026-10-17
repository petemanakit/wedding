/**
 * Pat & Pete RSVP endpoint.
 * Paste into the Google Sheet: Extensions > Apps Script, then Deploy > New deployment > Web app
 * (Execute as: Me; Who has access: Anyone). Copy the Web app URL into the website CONFIG.rsvpEndpoint.
 */
const SHEET_NAME = 'Responses';
const NOTIFY_EMAIL = 'peerakarn@ionenergy.ai';
const HEADERS = ['Timestamp', 'Invitation', 'Household / name', 'Guest names', 'Email', 'Phone',
  'Rehearsal dinner (Sat 16 Jan)', 'Vow ceremony (Sun 4.30 PM)', 'Evening celebration (Sun 6.30 PM)',
  'Headcount', 'Dietary requirements', 'Message', 'Kind'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME, 0); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#595D44').setFontColor('#F7F4EE');
    sh.setFrozenRows(1);
    ensureSummary_(ss);
  }
  return sh;
}

function ensureSummary_(ss) {
  if (ss.getSheetByName('Summary')) return;
  const s = ss.insertSheet('Summary');
  const r = "'" + SHEET_NAME + "'!";
  s.getRange('A1:C1').setValues([['Event', 'Accepting (households)', 'Guests (headcount)']]).setFontWeight('bold');
  s.getRange('A2:C4').setFormulas([
    ['Rehearsal dinner',   '=COUNTIF(' + r + 'G:G,"Accept")', '=SUMIF(' + r + 'G:G,"Accept",' + r + 'J:J)'],
    ['Vow ceremony',       '=COUNTIF(' + r + 'H:H,"Accept")', '=SUMIF(' + r + 'H:H,"Accept",' + r + 'J:J)'],
    ['Evening celebration','=COUNTIF(' + r + 'I:I,"Accept")', '=SUMIF(' + r + 'I:I,"Accept",' + r + 'J:J)'],
  ]);
  s.getRange('A6').setValue('Total replies'); s.getRange('B6').setFormula('=COUNTA(' + r + 'A:A)-1');
  s.getRange('A7').setValue('Declined all');   s.getRange('B7').setFormula('=COUNTIFS(' + r + 'I:I,"Decline")');
  s.autoResizeColumns(1, 3);
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return json_({ ok: true }); // honeypot filled: silently ignore
    const row = [
      new Date(), p.tier || '', p.household || '', p.guests || '', p.email || '', p.phone || '',
      p.rehearsal || '', p.vow || '', p.evening || '',
      Number(p.headcount || 0), p.dietary || '', p.message || '', p.kind || 'New'
    ];
    sheet_().appendRow(row);
    notify_(p);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return ContentService.createTextOutput('Pat & Pete RSVP endpoint is live.'); }

function notify_(p) {
  if (!NOTIFY_EMAIL) return;
  const lines = [
    'Household: ' + (p.household || ''), 'Guests: ' + (p.guests || ''), 'Headcount: ' + (p.headcount || ''),
    'Invitation: tier ' + (p.tier || ''), 'Rehearsal dinner: ' + (p.rehearsal || 'n/a'),
    'Vow ceremony: ' + (p.vow || 'n/a'), 'Evening: ' + (p.evening || ''),
    'Email: ' + (p.email || ''), 'Phone: ' + (p.phone || ''), 'Dietary: ' + (p.dietary || ''),
    'Message: ' + (p.message || ''), '', 'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ];
  const verdict = (p.evening === 'Accept' || p.vow === 'Accept' || p.rehearsal === 'Accept') ? 'accepts' : 'declines';
  MailApp.sendEmail(NOTIFY_EMAIL, 'RSVP: ' + (p.household || 'Guest') + ' ' + verdict, lines.join('\n'));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
