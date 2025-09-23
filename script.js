const TO_EMAIL = "you@example.com"; // энд өөрийн имэйл
const FILE_BASENAME = "ie_test_results";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents); // { meta, rows[] }
    const csv = buildCsv(data.rows);
    const when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");

    // CSV хавсралт
    const csvBlob = Utilities.newBlob(csv, "text/csv", `${FILE_BASENAME}_${when}.csv`);

    // XLSX үүсгэх (optional): эхлээд csv-гээ Google Sheet болгож, дараа нь xlsx болгон хувиргах
    const tempFolder = DriveApp.getRootFolder();
    const tempFile = DriveApp.createFile(`${FILE_BASENAME}_${when}.csv`, csv, MimeType.CSV);
    const ss = SpreadsheetApp.create(`${FILE_BASENAME}_${when}`);
    const sheet = ss.getActiveSheet();

    // CSV-г Sheet рүү буулгах
    const lines = csv.split("\n").filter(l => l.trim().length);
    const values = lines.map(line => line.split(",").map(s => s.replace(/^"|"$/g,"").replace(/""/g,'"')));
    sheet.getRange(1,1,values.length, values[0].length).setValues(values);

    // Sheet-ийг XLSX болгон хөрвүүлэх
    const xlsxBlob = DriveApp.getFileById(ss.getId()).getBlob().getAs(MimeType.MICROSOFT_EXCEL).setName(`${FILE_BASENAME}_${when}.xlsx`);

    // Имэйлийн доторхи тайлбар
    const summary = buildSummary(data);

    GmailApp.sendEmail(
      TO_EMAIL,
      `IE Test Results • ${when}`,
      summary,
      { attachments: [csvBlob, xlsxBlob] }
    );

    // Түр файлуудыг цэвэрлэх
    tempFile.setTrashed(true);
    DriveApp.getFileById(ss.getId()).setTrashed(true);

    return ContentService.createTextOutput("ok").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("error: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function buildCsv(rows) {
  const cols = ["idx","word","answer","rt_ms","ts","participant","valid","reason"];
  const header = cols.join(",");
  const lines = rows.map(r =>
    cols.map(c => {
      const v = (r[c] !== undefined && r[c] !== null) ? String(r[c]) : "";
      return `"${v.replace(/"/g,'""')}"`;
    }).join(",")
  );
  return header + "\n" + lines.join("\n");
}

function buildSummary(data) {
  const m = data.meta || {};
  const total = m.total ?? (data.rows?.length || 0);
  const valid = m.valid ?? 0;
  const invalid = total - valid;
  const avgI = isFinite(m.avg_rt_I) ? Math.round(m.avg_rt_I) : "-";
  const avgE = isFinite(m.avg_rt_E) ? Math.round(m.avg_rt_E) : "-";
  return [
    `Generated at: ${m.generated_at || ""}`,
    `Participant: ${m.participant || "-"}`,
    `Total: ${total}  •  Valid: ${valid}  •  Invalid: ${invalid}`,
    `Min valid RT: ${m.min_valid_rt_ms}ms`,
    `I avg RT: ${avgI} ms`,
    `E avg RT: ${avgE} ms`,
    `Spam window: ${m.same_key_spam_window} • Spam RT: ${m.same_key_spam_rt_ms}ms`
  ].join("\n");
}
