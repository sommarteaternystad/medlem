// =================================================================
// Fil: Code.gs
// Version 17.2: Tillägg av JSON-endpoints (getPasses, book) för att
// kunna anropas via fetch() från medlemssidan utan iframe.
// Inga befintliga funktioner är ändrade eller borttagna.
// =================================================================

const SPREADSHEET_ID = '1V4BY8WJpovQoIG9bcBqQ_vxxRzU6YZGcgXDPugwyg0w';
const SCHEMA_SHEET_NAME = 'Schema';
const BOOKINGS_SHEET_NAME = 'Bokningar';
const INFO_SHEET_NAME = 'InfoText';
const CONTACTS_SHEET_NAME = 'Ansvariga';

// --- HUVUDROUTER ---
function doGet(e) {
  // ── NYTT: JSON-endpoints för fristående frontend (medlemssidan) ──
  if (e.parameter.action === 'getPasses') {
    return ContentService.createTextOutput(JSON.stringify(getGroupedPasses()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e.parameter.action === 'book') {
    const result = bokaPass(e.parameter.passId, e.parameter.namn, e.parameter.epost, e.parameter.telefon);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e.parameter.action === 'cancel') {
    const template = HtmlService.createTemplateFromFile('cancel');
    template.passDetails = getPassDetailsForCancellation(e.parameter.bookingId, e.parameter.token);
    return template.evaluate()
      .setTitle('Avboka pass')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e.parameter.page === 'admin') {
    return HtmlService.createTemplateFromFile('admin')
      .evaluate()
      .setTitle('Admin - Cafébokning')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const template = HtmlService.createTemplateFromFile('index');
  template.infoData = getSystemData().infoText;
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Boka Cafépass')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- ANVÄNDARFUNKTIONER ---

/**
 * v17: Varje pass får eget kort (groupKey = passId).
 * Ansvarig hämtas per pass. Returnerar sorterad array med section-etikett.
 */
function getGroupedPasses() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const schemaSheet = ss.getSheetByName(SCHEMA_SHEET_NAME);
  const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);

  const passValues = schemaSheet.getDataRange().getValues();
  const passDisplayValues = schemaSheet.getDataRange().getDisplayValues();
  const allBookings = bookingsSheet.getDataRange().getDisplayValues();

  passValues.shift();
  passDisplayValues.shift();
  allBookings.shift();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const bookingCounts = {};
  const bookedByData = {};
  allBookings.forEach(booking => {
    const passId = booking[0];
    const status = booking[6];
    if (status === 'Bokad') {
      bookingCounts[passId] = (bookingCounts[passId] || 0) + 1;
      if (!bookedByData[passId]) bookedByData[passId] = [];
      bookedByData[passId].push({ name: booking[1] });
    }
  });

  const passes = [];

  for (let i = 0; i < passValues.length; i++) {
    const rawDate = passValues[i][1];

    if (rawDate instanceof Date && rawDate < today) continue;

    const t = passDisplayValues[i];
    const passId = t[0];
    const totalSlots = parseInt(t[5], 10) || 0;
    const currentBookings = bookingCounts[passId] || 0;

    passes.push({
      passId: passId,
      date: t[1],
      tid: t[2],
      event: t[3],
      uppgift: t[4],
      rawDate: rawDate instanceof Date ? rawDate.getTime() : 0,
      section: getSectionLabel(rawDate),
      huvudansvarig: t[6] || 'Ingen angiven',
      totalSlots: totalSlots,
      availableSlots: totalSlots - currentBookings,
      bookedBy: bookedByData[passId] || []
    });
  }

  // Sortera kronologiskt — fungerar alltid oavsett hur pass lagts in i kalkylark
  passes.sort((a, b) => a.rawDate - b.rawDate);
  return passes;
}

/**
 * Sektionsetikett baserat på månad — utbyggbar.
 * Maj–Jun  → Vårens föreställningar – Teaterskolan
 * Jul–Aug  → Sommarens föreställningar – Sommarteatern
 * Nov–Dec  → Jul hos Sommarteatern
 * Övrigt   → Kommande föreställningar
 */
function getSectionLabel(date) {
  if (!(date instanceof Date)) return 'Kommande föreställningar';
  const m = date.getMonth() + 1;
  if (m >= 5 && m <= 6)  return 'Vårens föreställningar – Teaterskolan';
  if (m >= 7 && m <= 8)  return 'Sommarens föreställningar – Sommarteatern';
  if (m >= 11 && m <= 12) return 'Jul hos Sommarteatern';
  return 'Kommande föreställningar';
}

function bokaPass(passId, namn, epost, telefon) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const schemaSheet = ss.getSheetByName(SCHEMA_SHEET_NAME);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);

    const passTemplates = schemaSheet.getDataRange().getValues();
    let passTemplate = null;
    for (let i = 1; i < passTemplates.length; i++) {
      if (String(passTemplates[i][0]).trim() == String(passId).trim()) {
        passTemplate = passTemplates[i]; break;
      }
    }
    if (!passTemplate) throw new Error("Passet kunde inte hittas.");

    const totalSlots = passTemplate[5];
    const allBookings = bookingsSheet.getDataRange().getValues();
    let currentBookings = 0;
    allBookings.forEach(b => {
      if (String(b[0]).trim() == String(passId).trim() && b[6] === 'Bokad') currentBookings++;
    });

    if (currentBookings >= totalSlots) {
      return { success: false, message: 'Tyvärr hann den sista platsen bokas.' };
    }

    const token = Utilities.getUuid();
    bookingsSheet.appendRow([passId, namn, epost, telefon, new Date(), token, 'Bokad', '']);

    const dateObj = new Date(passTemplate[1]);
    const passInfo = {
      namn, epost,
      datum: dateObj.toLocaleDateString('sv-SE', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }),
      datumForParsing: passTemplate[1],
      tid: passTemplate[2],
      forestallning: passTemplate[3],
      uppgift: passTemplate[4],
      bookingId: bookingsSheet.getLastRow(),
      token,
      huvudansvarig: passTemplate[6]
    };
    sendConfirmationEmail(passInfo);

    return { success: true, message: 'Tack! Passet är bokat – kolla din e-post.' };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: `Ett tekniskt fel uppstod: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function avbokaPass(params) {
  const { bookingId, token, reason } = params;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
  const schemaSheet = ss.getSheetByName(SCHEMA_SHEET_NAME);

  try {
    const bookingData = bookingsSheet.getRange(bookingId, 1, 1, 8).getValues()[0];
    if (bookingData[5] !== token) return "Avbokningen misslyckades. Länken är ogiltig.";

    const passId = bookingData[0];
    const passTemplates = schemaSheet.getDataRange().getDisplayValues();
    let passInfo = { datum: 'Okänt', tid: 'Okänt', forestallning: 'Okänt', uppgift: '', huvudansvarig: '' };
    for (let i = 1; i < passTemplates.length; i++) {
      if (passTemplates[i][0] == passId) {
        passInfo = { datum: passTemplates[i][1], tid: passTemplates[i][2], forestallning: passTemplates[i][3], uppgift: passTemplates[i][4], huvudansvarig: passTemplates[i][6] };
        break;
      }
    }

    bookingsSheet.getRange(bookingId, 7, 1, 2).setValues([['Avbokad', reason]]);
    sendCancellationEmail(bookingData[1], bookingData[2], passInfo, reason);
    return "Ditt pass har avbokats. En bekräftelse har skickats till din e-post.";
  } catch (e) {
    Logger.log(e);
    return "Ett tekniskt fel uppstod.";
  }
}

// --- ADMIN-FUNKTIONER ---
function verifyAdminAndGetData(initials, code) {
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONTACTS_SHEET_NAME).getDataRange().getValues();
  let isValid = false, adminFullName = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toLowerCase().trim() === initials.toLowerCase().trim() && data[i][3] === code) {
      isValid = true; adminFullName = data[i][0]; break;
    }
  }
  if (isValid) {
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put(token, adminFullName, 3600);
    return { success: true, data: fetchAllPassesAndStats(), token, adminName: adminFullName };
  }
  return { success: false, message: 'Fel initialer eller kod.' };
}

function refreshAdminBookings(token) {
  if (CacheService.getScriptCache().get(token)) return fetchAllPassesAndStats();
  return { stats: { total: 0, booked: 0, available: 0 }, passes: {}, bookings: [] };
}

function adminDeleteBooking(bookingRowId, token) {
  const adminName = CacheService.getScriptCache().get(token);
  if (!adminName) return { success: false, message: 'Din session har gått ut.' };
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BOOKINGS_SHEET_NAME)
      .getRange(bookingRowId, 7, 1, 2).setValues([['Raderad av admin', `Raderad av ${adminName}`]]);
    return { success: true, message: 'Bokningen har raderats.' };
  } catch (e) {
    return { success: false, message: 'Ett fel uppstod.' };
  }
}

/**
 * Admin bokar in en person manuellt — namn krävs, e-post och telefon valfria.
 * Skickar bekräftelsemail bara om e-post finns.
 */
function adminBokaManually(passId, namn, epost, telefon, token) {
  const adminName = CacheService.getScriptCache().get(token);
  if (!adminName) return { success: false, message: 'Din session har gått ut.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const schemaSheet = ss.getSheetByName(SCHEMA_SHEET_NAME);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);

    const passTemplates = schemaSheet.getDataRange().getValues();
    let passTemplate = null;
    for (let i = 1; i < passTemplates.length; i++) {
      if (String(passTemplates[i][0]).trim() == String(passId).trim()) {
        passTemplate = passTemplates[i]; break;
      }
    }
    if (!passTemplate) return { success: false, message: 'Passet hittades inte.' };

    const totalSlots = passTemplate[5];
    const allBookings = bookingsSheet.getDataRange().getValues();
    let currentBookings = 0;
    allBookings.forEach(b => {
      if (String(b[0]).trim() == String(passId).trim() && b[6] === 'Bokad') currentBookings++;
    });

    if (currentBookings >= totalSlots) {
      return { success: false, message: 'Passet är fullbokat.' };
    }

    const bookingToken = Utilities.getUuid();
    bookingsSheet.appendRow([passId, namn, epost || '', telefon || '', new Date(), bookingToken, 'Bokad', 'Inlagd av admin (' + adminName + ')']);

    // Skicka bekräftelsemail bara om e-post finns
    if (epost) {
      const dateObj = new Date(passTemplate[1]);
      const passInfo = {
        namn: namn, epost: epost,
        datum: dateObj.toLocaleDateString('sv-SE', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }),
        datumForParsing: passTemplate[1],
        tid: passTemplate[2],
        forestallning: passTemplate[3],
        uppgift: passTemplate[4],
        bookingId: bookingsSheet.getLastRow(),
        token: bookingToken,
        huvudansvarig: passTemplate[6]
      };
      sendConfirmationEmail(passInfo);
    }

    return { success: true, message: namn + ' är inbokad!' + (epost ? ' Bekräftelsemail skickat.' : '') };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: 'Fel: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function adminUpdateHuvudansvarig(passId, nyAnsvarig, token) {
  if (!CacheService.getScriptCache().get(token)) return { success: false, message: 'Din session har gått ut.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SCHEMA_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(passId).trim()) {
        sheet.getRange(i + 1, 7).setValue(nyAnsvarig);
        return { success: true, message: `Ansvarig uppdaterad till ${nyAnsvarig}.` };
      }
    }
    return { success: false, message: 'Passet hittades inte.' };
  } catch (e) {
    return { success: false, message: 'Ett tekniskt fel uppstod.' };
  }
}

function fetchAllPassesAndStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rawRows = ss.getSheetByName(SCHEMA_SHEET_NAME).getDataRange().getValues();
  const dispRows = ss.getSheetByName(SCHEMA_SHEET_NAME).getDataRange().getDisplayValues();
  const bookDisp = ss.getSheetByName(BOOKINGS_SHEET_NAME).getDataRange().getDisplayValues();

  rawRows.shift(); dispRows.shift(); bookDisp.shift();

  const passesById = {};
  for (let i = 0; i < dispRows.length; i++) {
    const t = dispRows[i];
    const rawDate = rawRows[i][1];
    passesById[t[0]] = {
      datum: t[1], tid: t[2], forestallning: t[3], uppgift: t[4],
      total: parseInt(t[5] || '0', 10),
      huvudansvarig: t[6] || '',
      rawDate: rawDate instanceof Date ? rawDate.getTime() : 0
    };
  }

  let bookedCount = 0;
  bookDisp.forEach(b => { if (b[6] === 'Bokad') bookedCount++; });
  const total = dispRows.reduce((s, t) => s + (parseInt(t[5] || '0', 10)), 0);

  const bookings = bookDisp.map((b, i) => ({
    rowId: i + 2, passId: b[0], namn: b[1], epost: b[2],
    telefon: formatPhone(b[3]), bokadDatum: b[4], token: b[5], status: b[6], anledning: b[7]
  }));

  const ansvariga = ss.getSheetByName(CONTACTS_SHEET_NAME).getDataRange().getValues()
    .slice(1).map(r => r[0]).filter(n => n);

  return { stats: { total, booked: bookedCount, available: total - bookedCount }, bookings, passes: passesById, ansvariga };
}

// --- HJÄLPFUNKTIONER ---
function getSystemData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const infoData = ss.getSheetByName(INFO_SHEET_NAME).getRange("A2:B2").getValues()[0];
  const infoText = { title: infoData[0], content: infoData[1].replace(/\n/g, '<br>') };
  const contacts = {};
  ss.getSheetByName(CONTACTS_SHEET_NAME).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) contacts[r[0].trim()] = r[1];
  });
  return { infoText, contacts };
}

function getPassDetailsForCancellation(bookingId, token) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const bookingData = ss.getSheetByName(BOOKINGS_SHEET_NAME).getRange(bookingId, 1, 1, 8).getValues()[0];
    if (bookingData[5] !== token || bookingData[6] !== 'Bokad') return null;
    const passId = bookingData[0];
    const rows = ss.getSheetByName(SCHEMA_SHEET_NAME).getDataRange().getDisplayValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == passId) return { bookingId, token, datum: rows[i][1], tid: rows[i][2], forestallning: rows[i][3] };
    }
    return null;
  } catch (e) { return null; }
}

// --- E-POST ---

function sendConfirmationEmail(passInfo) {
  const systemData = getSystemData();
  const dateObject = passInfo.datumForParsing;
  const ds = dateObject.getFullYear() + '-' + (dateObject.getMonth() + 1) + '-' + dateObject.getDate();
  const startTime = parseDateTime(ds, passInfo.tid);
  if (!startTime) throw new Error("Could not create a valid date object.");

  let durationH = 2;
  const tidClean = passInfo.tid.replace(/\s/g, '');
  if (tidClean.includes('-')) {
    try {
      const parts = tidClean.split('-');
      const s = parseDateTime(ds, parts[0]), en = parseDateTime(ds, parts[1]);
      if (s && en && en > s) durationH = (en - s) / 3600000;
    } catch (e) {}
  }
  const endTime = new Date(startTime.getTime() + durationH * 3600000);

  const ansvarigNamn = passInfo.huvudansvarig || 'Ej angiven';
  const ansvarigTel = systemData.contacts[ansvarigNamn.trim()] || 'Saknas';
  const webAppUrl = ScriptApp.getService().getUrl();
  const fmt = d => d.toISOString().replace(/-|:|\.\d{3}/g, '');
  const calLink = `https://www.google.com/calendar/render?action=TEMPLATE&text=${encodeURIComponent('Arbetspass: ' + passInfo.forestallning)}&dates=${fmt(startTime)}/${fmt(endTime)}&details=${encodeURIComponent('Uppgift: ' + passInfo.uppgift)}`;
  const cancelLink = `${webAppUrl}?action=cancel&bookingId=${passInfo.bookingId}&token=${passInfo.token}`;
  const subject = `✅ Bokningsbekräftelse – ${passInfo.forestallning} ${passInfo.datum}`;

  const html = `<!DOCTYPE html><html lang="sv"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f2f2f2;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2f2;padding:40px 16px;">
<tr><td align="center">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 4px 32px rgba(0,0,0,0.10);">

  <tr><td style="background:linear-gradient(140deg,#14532d 0%,#16a34a 100%);padding:40px;text-align:center;">
    <p style="margin:0 0 12px;font-size:40px;">🎭</p>
    <h1 style="margin:0;color:#fff;font-size:24px;font-weight:800;letter-spacing:-0.5px;">Bokningsbekräftelse</h1>
    <p style="margin:10px 0 0;color:rgba(255,255,255,0.75);font-size:14px;">Tack ${passInfo.namn} – vi ser fram emot att se dig!</p>
  </td></tr>

  <tr><td style="padding:36px 40px 0;">
    <p style="margin:0 0 4px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:0.1em;color:#16a34a;">DITT PASS</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1.5px solid #e8e8e8;border-radius:12px;overflow:hidden;margin-top:12px;">
      <tr style="background:#fafafa;"><td style="padding:14px 18px;border-bottom:1px solid #e8e8e8;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Föreställning</span>
        <div style="font-size:16px;font-weight:700;color:#111;margin-top:4px;">${passInfo.forestallning}</div>
      </td></tr>
      <tr><td style="padding:14px 18px;border-bottom:1px solid #e8e8e8;background:#fff;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Datum</span>
        <div style="font-size:16px;font-weight:700;color:#111;margin-top:4px;">${passInfo.datum}</div>
      </td></tr>
      <tr style="background:#fafafa;"><td style="padding:14px 18px;border-bottom:1px solid #e8e8e8;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Tid</span>
        <div style="font-size:16px;font-weight:700;color:#111;margin-top:4px;">${passInfo.tid}</div>
      </td></tr>
      <tr><td style="padding:14px 18px;background:#fff;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Din uppgift</span>
        <div style="font-size:16px;font-weight:700;color:#111;margin-top:4px;">${passInfo.uppgift}</div>
      </td></tr>
    </table>
  </td></tr>

  <tr><td style="padding:24px 40px 0;">
    <div style="background:#f0fdf4;border:1.5px solid #bbf7d0;border-radius:12px;padding:18px 20px;">
      <p style="margin:0 0 3px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:0.1em;color:#16a34a;">Huvudansvarig för passet</p>
      <p style="margin:0;font-size:17px;font-weight:800;color:#111;">${ansvarigNamn}</p>
      <p style="margin:5px 0 0;font-size:14px;color:#555;">📞 ${ansvarigTel}</p>
      <p style="margin:8px 0 0;font-size:12px;color:#888;font-style:italic;">Vid kort varsel – kontakta ansvarig direkt istället för att avboka online.</p>
    </div>
  </td></tr>

  <tr><td style="padding:24px 40px 0;text-align:center;">
    <a href="${calLink}" style="display:inline-block;background:#111;color:#fff;text-decoration:none;padding:14px 30px;border-radius:10px;font-size:14px;font-weight:700;letter-spacing:0.02em;">📅 Lägg till i Google Calendar</a>
  </td></tr>

  <tr><td style="padding:24px 40px 0;">
    <div style="border:1.5px solid #e8e8e8;border-radius:12px;padding:18px 20px;">
      <p style="margin:0 0 6px;font-size:14px;font-weight:700;color:#111;">Behöver du avboka?</p>
      <p style="margin:0 0 14px;font-size:13px;color:#777;line-height:1.6;">Gör det i god tid via länken nedan. Ring ansvarig vid kort varsel.</p>
      <a href="${cancelLink}" style="display:inline-block;background:#fff;border:1.5px solid #fca5a5;color:#dc2626;text-decoration:none;padding:10px 20px;border-radius:8px;font-size:13px;font-weight:700;">Avboka mitt pass</a>
    </div>
  </td></tr>

  <tr><td style="padding:24px 40px 0;">
    <div style="background:#fffbeb;border:1.5px solid #fde68a;border-radius:12px;padding:18px 20px;">
      <p style="margin:0 0 6px;font-size:13px;font-weight:700;color:#92400e;">${systemData.infoText.title}</p>
      <p style="margin:0;font-size:13px;color:#78350f;line-height:1.7;">${systemData.infoText.content}</p>
    </div>
  </td></tr>

  <tr><td style="padding:36px 40px;text-align:center;">
    <p style="margin:0;font-size:12px;color:#ccc;">Sommarteatern &nbsp;·&nbsp; Arbetspass-bokning</p>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>`;

  MailApp.sendEmail({ to: passInfo.epost, subject, htmlBody: html });
}

function sendCancellationEmail(namn, epost, passInfo, reason) {
  const subject = `Avbokning bekräftad – ${passInfo.forestallning} ${passInfo.datum}`;

  const html = `<!DOCTYPE html><html lang="sv"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f2f2f2;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f2f2f2;padding:40px 16px;">
<tr><td align="center">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 4px 32px rgba(0,0,0,0.10);">

  <tr><td style="background:linear-gradient(140deg,#7f1d1d 0%,#ef4444 100%);padding:40px;text-align:center;">
    <p style="margin:0 0 12px;font-size:40px;">📋</p>
    <h1 style="margin:0;color:#fff;font-size:24px;font-weight:800;">Avbokning bekräftad</h1>
    <p style="margin:10px 0 0;color:rgba(255,255,255,0.75);font-size:14px;">Hej ${namn} – din avbokning är registrerad.</p>
  </td></tr>

  <tr><td style="padding:36px 40px 0;">
    <p style="margin:0 0 4px;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:0.1em;color:#dc2626;">AVBOKAT PASS</p>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1.5px solid #e8e8e8;border-radius:12px;overflow:hidden;margin-top:12px;">
      <tr style="background:#fafafa;"><td style="padding:14px 18px;border-bottom:1px solid #e8e8e8;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Föreställning</span>
        <div style="font-size:15px;font-weight:700;color:#111;margin-top:4px;">${passInfo.forestallning}</div>
      </td></tr>
      <tr><td style="padding:14px 18px;border-bottom:1px solid #e8e8e8;background:#fff;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Datum</span>
        <div style="font-size:15px;font-weight:700;color:#111;margin-top:4px;">${passInfo.datum}</div>
      </td></tr>
      <tr style="background:#fafafa;"><td style="padding:14px 18px;border-bottom:1px solid #e8e8e8;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Tid</span>
        <div style="font-size:15px;font-weight:700;color:#111;margin-top:4px;">${passInfo.tid}</div>
      </td></tr>
      <tr><td style="padding:14px 18px;background:#fff;">
        <span style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.07em;color:#888;">Angiven anledning</span>
        <div style="font-size:15px;font-weight:600;color:#555;margin-top:4px;font-style:italic;">"${reason}"</div>
      </td></tr>
    </table>
  </td></tr>

  <tr><td style="padding:24px 40px 0;">
    <div style="background:#fafafa;border:1.5px solid #e8e8e8;border-radius:12px;padding:18px 20px;text-align:center;">
      <p style="margin:0;font-size:14px;color:#555;line-height:1.7;">Vill du boka ett annat pass? Gå tillbaka till bokningssidan och välj ett nytt datum.<br>Vi hoppas vi ses snart! 🎭</p>
    </div>
  </td></tr>

  <tr><td style="padding:36px 40px;text-align:center;">
    <p style="margin:0;font-size:12px;color:#ccc;">Sommarteatern &nbsp;·&nbsp; Arbetspass-bokning</p>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>`;

  MailApp.sendEmail({ to: epost, subject, htmlBody: html });
}

/**
 * Säkerställer att telefonnummer behåller sin ledande nolla.
 * Google Sheets lagrar telefonnummer som siffror och tappar inledande 0.
 */
function formatPhone(val) {
  if (!val) return '';
  let s = String(val).trim();
  if (/^\d{9}$/.test(s)) s = '0' + s; // 9 siffror → lägg till ledande 0
  return s;
}

function parseDateTime(dateString, timeString) {
  try {
    const dd = dateString.match(/\d+/g);
    if (!dd || dd.length < 3) return null;
    const td = timeString.match(/\d+/g);
    if (!td) return null;
    const d = new Date(parseInt(dd[0]), parseInt(dd[1]) - 1, parseInt(dd[2]), parseInt(td[0]), td.length > 1 ? parseInt(td[1]) : 0);
    return isNaN(d.getTime()) ? null : d;
  } catch (e) { return null; }
}
