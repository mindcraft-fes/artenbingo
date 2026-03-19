// ── ARTENBINGO RATINGEN – Web App ──
// 1. Öffne dein Google Sheet
// 2. Erweiterungen → Apps Script
// 3. Diesen Code einfügen → Speichern
// 4. Bereitstellen → Neue Bereitstellung → Web-App
//    - Ausführen als: Ich
//    - Zugriff: Jeder
// 5. Die Web-App-URL ins Bingo eintragen (WEBAPP_URL)

const SHEET_NAME = "Ergebnisse";

// ── GET: Ranking-Daten zurückgeben ──
function doGet(e) {
  // CORS-Header setzen
  const action = e && e.parameter && e.parameter.action;

  if (action === 'ranking') {
    try {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const blatt = ss.getSheetByName(SHEET_NAME);
      if (!blatt || blatt.getLastRow() < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ status:"ok", rows:[] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const lastRow = blatt.getLastRow();
      const data    = blatt.getRange(2, 1, lastRow-1, 9).getValues();
      const rows = data
        .filter(r => r[1] !== '')
        .map(r => ({
          timestamp:   r[0] ? String(r[0]) : '',
          name:        String(r[1]),
          klasse:      String(r[2]),
          pflanzen:    Number(r[3]) || 0,
          baeume:      Number(r[4]) || 0,
          voegel:      Number(r[5]) || 0,
          insekten:    Number(r[6]) || 0,
          saeugetiere: Number(r[7]) || 0,
          gesamt:      Number(r[8]) || 0,
        }));
      return ContentService
        .createTextOutput(JSON.stringify({ status:"ok", rows }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status:"error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createHtmlOutput("🌿 Artenbingo Ratingen – Web-App läuft!");
}

// ── POST: Empfängt Daten vom Bingo ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    speichereErgebnis(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Ergebnis ins Sheet schreiben ──
function speichereErgebnis(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let blatt = ss.getSheetByName(SHEET_NAME);

  // Blatt anlegen falls nicht vorhanden
  if (!blatt) {
    blatt = ss.insertSheet(SHEET_NAME);
    const header = ["Zeitstempel", "Name", "Klasse", "Pflanzen", "Bäume & Sträucher", "Vögel", "Insekten", "Säugetiere", "Gesamt"];
    blatt.getRange(1, 1, 1, header.length).setValues([header]);
    const hr = blatt.getRange(1, 1, 1, header.length);
    hr.setBackground("#1a3d1a").setFontColor("white").setFontWeight("bold").setFontSize(11);
    blatt.setFrozenRows(1);
    blatt.setColumnWidth(1, 160);
    blatt.setColumnWidth(2, 160);
    blatt.setColumnWidth(3, 80);
    for (let i = 4; i <= 9; i++) blatt.setColumnWidth(i, 110);
  }

  // Prüfen ob dieser Schüler schon eingetragen ist → Zeile aktualisieren
  const lastRow = blatt.getLastRow();
  let gefunden  = -1;
  if (lastRow > 1) {
    const namen = blatt.getRange(2, 2, lastRow - 1, 2).getValues();
    namen.forEach((row, i) => {
      if (row[0] === data.name && row[1] === data.klasse) gefunden = i + 2;
    });
  }

  const zeile = [
    new Date().toLocaleString("de-DE"),
    data.name,
    data.klasse,
    data.pflanzen    || 0,
    data.baeume      || 0,
    data.voegel      || 0,
    data.insekten    || 0,
    data.saeugetiere || 0,
    data.gesamt      || 0
  ];

  if (gefunden > 0) {
    blatt.getRange(gefunden, 1, 1, zeile.length).setValues([zeile]);
  } else {
    blatt.appendRow(zeile);
  }

  aktualisiereRanking();
}

// ── RANKING AKTUALISIEREN ──
function aktualisiereRanking() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const daten = ss.getSheetByName(SHEET_NAME);
  if (!daten || daten.getLastRow() < 2) return;

  let ranking = ss.getSheetByName("Ranking");
  if (!ranking) ranking = ss.insertSheet("Ranking");
  ranking.clearContents().clearFormats();

  // Titel
  ranking.getRange("A1").setValue("🏆 Artenbingo Ratingen – Ranking");
  ranking.getRange("A1").setFontSize(16).setFontWeight("bold").setFontColor("#1a3d1a");
  ranking.getRange("A1:J1").merge();
  ranking.getRange("A2").setValue("Stand: " + new Date().toLocaleString("de-DE"));
  ranking.getRange("A2").setFontColor("#888").setFontSize(10);
  ranking.getRange("A2:J2").merge();

  // Header
  const header = ["Platz", "Name", "Klasse", "🌿 Pflanzen", "🌳 Bäume", "🐦 Vögel", "🐝 Insekten", "🦔 Säugetiere", "⭐ Gesamt", "Zuletzt gesendet"];
  ranking.getRange(3, 1, 1, header.length).setValues([header]);
  ranking.getRange(3, 1, 1, header.length)
    .setBackground("#1a3d1a").setFontColor("white").setFontWeight("bold").setFontSize(11);
  ranking.setFrozenRows(3);

  // Daten holen & sortieren
  const lastRow  = daten.getLastRow();
  const rohdaten = daten.getRange(2, 1, lastRow - 1, 9).getValues().filter(r => r[1] !== "");
  rohdaten.sort((a, b) => Number(b[8]) - Number(a[8]));

  rohdaten.forEach((row, i) => {
    const zeile = i + 4;
    const platz = i + 1;
    let platzText = String(platz);
    if (platz === 1) platzText = "🥇 1";
    if (platz === 2) platzText = "🥈 2";
    if (platz === 3) platzText = "🥉 3";

    ranking.getRange(zeile, 1, 1, 10).setValues([[
      platzText, row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[0]
    ]]);

    const bg = platz === 1 ? "#fff8dc" : platz === 2 ? "#f0f0f0" : platz === 3 ? "#fff0e0" : i % 2 === 0 ? "#f5f0e8" : "#ffffff";
    ranking.getRange(zeile, 1, 1, 10).setBackground(bg);
    if (platz <= 3) ranking.getRange(zeile, 1, 1, 10).setFontWeight("bold");
    ranking.getRange(zeile, 9).setFontWeight("bold").setFontColor("#1a3d1a");
  });

  // Spaltenbreiten
  ranking.setColumnWidth(1, 75);
  ranking.setColumnWidth(2, 160);
  ranking.setColumnWidth(3, 75);
  for (let i = 4; i <= 9; i++) ranking.setColumnWidth(i, 95);
  ranking.setColumnWidth(10, 150);

  // Statistik
  if (rohdaten.length > 0) {
    const statsZ = rohdaten.length + 6;
    ranking.getRange(statsZ, 1).setValue("📊 Statistik").setFontWeight("bold").setFontSize(12).setFontColor("#1a3d1a");
    ranking.getRange(statsZ+1, 1).setValue("Teilnehmer:");
    ranking.getRange(statsZ+1, 2).setValue(rohdaten.length);
    const gesamt = rohdaten.map(r => Number(r[8]));
    ranking.getRange(statsZ+2, 1).setValue("Höchste Punktzahl:");
    ranking.getRange(statsZ+2, 2).setValue(Math.max(...gesamt));
    ranking.getRange(statsZ+3, 1).setValue("Durchschnitt:");
    ranking.getRange(statsZ+3, 2).setValue((gesamt.reduce((a,b)=>a+b,0)/gesamt.length).toFixed(1));
  }
}

// ── MENÜ ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🌿 Artenbingo")
    .addItem("🏆 Ranking jetzt aktualisieren", "aktualisiereRanking")
    .addToUi();
}
