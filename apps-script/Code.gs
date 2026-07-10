function doGet(e) {
  // ── NYTT: JSON-endpoint för fristående frontend (medlemssidan) ──
  if (e && e.parameter && e.parameter.action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(getData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PRODUKTIONSSCHEMA 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName("Blad 1");
  const mainData = sheet1.getDataRange().getValues().slice(1).map(row => {
    if (!row[0]) return null;
    let d = new Date(row[0]);
    let isoDate = Utilities.formatDate(d, "Europe/Stockholm", "yyyy-MM-dd");
    return {
      date: isoDate,
      time: String(row[1]),
      title: String(row[2]),
      desc: String(row[3]),
      cat: String(row[4]).trim().toUpperCase(),
      link: String(row[5] || "")
    };
  }).filter(r => r !== null);

  const sheet2 = ss.getSheetByName("Blad 2");
  const detailData = sheet2.getDataRange().getValues().slice(1).map(row => {
    if (!row[0]) return null;
    let d = new Date(row[0]);
    let isoDate = Utilities.formatDate(d, "Europe/Stockholm", "yyyy-MM-dd");
    return {
      date: isoDate,
      time: String(row[1]),
      title: String(row[2]),
      desc: String(row[3]),
      who: String(row[4]),
      place: String(row[5] || ""),
      category: String(row[6] || "").trim().toLowerCase()
    };
  }).filter(r => r !== null);

  return { main: mainData, details: detailData };
}
