function getSheetRow(sheet, row) {
  // Google 시트 가져오기
  var sheetData = sheet.getDataRange().getValues();

  // 특정 행 데이터 추출
  var rowData = sheetData[row - 1];

  // Google Docs 본문에 텍스트 추가
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  body.appendParagraph(rowData.join('\n'));
}

// 예시: 2행 내용을 Google Docs로 가져오기
var sheet =SpreadsheetApp.openById('11bldxPQ7agyw_tqnSS4ywyNkIzQRHCI94s16Vbf3txI');
var row = 2;
getSheetRow(sheet, row);
