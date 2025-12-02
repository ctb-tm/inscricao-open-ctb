function doGet(e) {
  const sheet = SpreadsheetApp.openById("1m_ai30cbLA79RgKrjx55sqPWOSLNRPDUbwQUnqCPBzM")
                .getSheetByName("Respostas");

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // primeira linha

  const json = data.map(row => {
    const item = {};
    row.forEach((value, i) => item[headers[i]] = value);
    return item;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", data: json }))
    .setMimeType(ContentService.MimeType.JSON);
}
