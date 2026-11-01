// ===== Instellingen =====
const TRAINER = 'KVBU7';              // geheime trainerscode
const FEEDS = {                               // de .ics-links van Foot24 (iCal-link kopiëren)
  A: 'https://www.foot24.be/nl/ical/kv-bonheiden/u7-a-1141.ics?key=m71PmnR',
  B: 'https://www.foot24.be/nl/ical/kv-bonheiden/u7-b-1343.ics?key=nG2PnWR'
};
const TZ = 'Europe/Brussels';

// ===== Hulpfuncties =====
function sh_(n, kop) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(kop); }
  return s;
}
const sp_ = () => sh_('Spelers', ['naam']);
const an_ = () => sh_('Antwoorden', ['wedstrijd', 'speler', 'antwoord', 'opmerking', 'tijd']);
const in_ = () => sh_('Indeling', ['wedstrijd', 'ploeg', 'datum', 'spelers', 'bevestigd']);
const ex_ = () => sh_('Extra wedstrijden', ['ploeg (A of B)', 'datum (jjjj-mm-dd)', 'uur (uu:mm)', 'tegenstander', 'thuis (ja/nee)', 'plaats']);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const rows_ = s => s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues() : [];
function fmt_(v, f) {
  if (Object.prototype.toString.call(v) === '[object Date]')
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), f || 'yyyy-MM-dd');
  return String(v);
}

// ===== Wedstrijden uit de .ics =====
function parseIcs_(txt, ploeg) {
  const res = [];
  const t = txt.replace(/\r?\n[ \t]/g, '');
  t.split('BEGIN:VEVENT').slice(1).forEach(b => {
    const g = k => {
      const m = b.match(new RegExp('^' + k + '[^:\\r\\n]*:(.*)$', 'm'));
      return m ? m[1].trim().replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\n/gi, ' ') : '';
    };
    const d = g('DTSTART');
    const m = d.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2}))?/);
    if (!m) return;
    let datum = m[1] + '-' + m[2] + '-' + m[3], uur = m[4] ? m[4] + ':' + m[5] : '';
    if (/Z$/.test(d)) {
      const dt = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)));
      datum = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
      uur = Utilities.formatDate(dt, TZ, 'HH:mm');
    }
    const s = g('SUMMARY');
    const parts = s.split(/\s+(?:-|–|vs\.?|tegen)\s+/i);
    let thuis = true, tegen = s;
    if (parts.length >= 2) {
      const i = parts.findIndex(p => /bonheiden/i.test(p));
      thuis = i !== 1;
      tegen = parts[i === 0 ? 1 : 0];
    }
    res.push({ id: ploeg + '|' + (g('UID') || datum + uur), ploeg, datum, uur, tegen, thuis, plaats: g('LOCATION') });
  });
  return res;
}

function extra_() {
  return rows_(ex_()).filter(r => r[1]).map((r, i) => ({
    id: 'X|' + String(r[0]).trim().toUpperCase() + '|' + fmt_(r[1]) + '|' + fmt_(r[2], 'HH:mm'),
    ploeg: String(r[0]).trim().toUpperCase(), datum: fmt_(r[1]), uur: fmt_(r[2], 'HH:mm'),
    tegen: String(r[3]), thuis: String(r[4]).toLowerCase() !== 'nee', plaats: String(r[5])
  }));
}

function wedstrijden_() {
  const c = CacheService.getScriptCache(), hit = c.get('wed');
  let all;
  if (hit) all = JSON.parse(hit);
  else {
    all = [];
    Object.keys(FEEDS).forEach(p => {
      try { all = all.concat(parseIcs_(UrlFetchApp.fetch(FEEDS[p], { muteHttpExceptions: true }).getContentText(), p)); } catch (e) {}
    });
    if (all.length) c.put('wed', JSON.stringify(all), 1800);
  }
  const vandaag = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return all.concat(extra_()).filter(w => w.datum >= vandaag).sort((a, b) => (a.datum + a.uur).localeCompare(b.datum + b.uur));
}

// ===== Web-app =====
function doGet(e) {
  const p = e.parameter, trainer = p.code === TRAINER;
  const spelers = rows_(sp_()).map(r => String(r[0])).filter(String);
  const ind = rows_(in_()).map(r => ({
    id: String(r[0]), ploeg: r[1], datum: fmt_(r[2]),
    spelers: String(r[3]).split('|').filter(String), bevestigd: r[4] === true || r[4] === 'TRUE'
  }));
  const an = rows_(an_()).filter(r => trainer || String(r[1]) === p.speler)
    .map(r => ({ id: String(r[0]), speler: String(r[1]), antwoord: r[2], opm: r[3] }));
  return out_({ wedstrijden: wedstrijden_(), spelers, antwoorden: an, indeling: trainer ? ind : ind.filter(i => i.bevestigd), trainer });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents), trainer = d.code === TRAINER;
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (d.actie === 'antwoord') {
      const s = an_(), v = s.getDataRange().getValues();
      const i = v.findIndex((r, k) => k > 0 && String(r[0]) === d.id && String(r[1]) === d.speler);
      const row = [d.id, d.speler, d.antwoord, d.opm || '', new Date()];
      if (i > 0) s.getRange(i + 1, 1, 1, 5).setValues([row]); else s.appendRow(row);
    } else if (!trainer) {
      return out_({ error: 'Alleen voor de trainer' });
    } else if (d.actie === 'indeling') {
      const s = in_(), v = s.getDataRange().getValues();
      const i = v.findIndex((r, k) => k > 0 && String(r[0]) === d.id);
      s.getRange('C:C').setNumberFormat('@');
      const row = [d.id, d.ploeg, d.datum, d.spelers.join('|'), !!d.bevestigd];
      if (i > 0) s.getRange(i + 1, 1, 1, 5).setValues([row]); else s.appendRow(row);
    } else if (d.actie === 'spelers') {
      const s = sp_();
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 1).clearContent();
      if (d.lijst.length) s.getRange(2, 1, d.lijst.length, 1).setValues(d.lijst.map(n => [n]));
    }
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}
