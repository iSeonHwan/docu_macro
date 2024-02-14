/* 구글 시트에서 구글 독스로 자료 가져오기: 특정 열 정보 가져오기 */

function getSheetColumn(sheet, column) {
  // Google 시트 가져오기
  var sheetData = sheet.getDataRange().getValues();

  // 특정 열 데이터 추출
  var columnData = sheetData.map(function(row) {
    return row[column - 1];
  });

  // Google Docs 본문에 텍스트 추가
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  body.appendParagraph(columnData.join('\n'));
}

// 예시: A열 내용을 Google Docs로 가져오기
var sheet =SpreadsheetApp.openById('11bldxPQ7agyw_tqnSS4ywyNkIzQRHCI94s16Vbf3txI');
var column = 1;
getSheetColumn(sheet, column);
