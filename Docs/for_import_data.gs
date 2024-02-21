function importSheetData() {
  // 시트 ID 및 범위 입력
  var sheetId = '11bldxPQ7agyw_tqnSS4ywyNkIzQRHCI94s16Vbf3txI';
  var range = 'A1:F4';

  // 시트 열기
  var sheet = SpreadsheetApp.openById(sheetId);

  // 문서 열기
  var doc = DocumentApp.getActiveDocument();

  // 데이터 가져오기
  var data = sheet.getRange(range).getValues();

  // 데이터 추가
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      doc.appendParagraph(data[i][j]);
    }
  }
}
