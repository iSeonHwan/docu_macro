function ImportDataFromSheet() {
  
  // Google Sheets 데이터 가져오기
  // 메서드를 사용하여 스프레드시트 ID를 기준으로 스프레드시트를 엽니다.
  var sheet = SpreadsheetApp.openById('11bldxPQ7agyw_tqnSS4ywyNkIzQRHCI94s16Vbf3txI');

  //메서드를 사용하여 데이터를 가져올 범위를 선택합니다.
  var range = sheet.getRange('B2:B19');

  //메서드를 사용하여 범위의 값을 배열로 가져옵니다.
  var values = range.getValues();

  // Google Docs 본문에 데이터 추가
  // 메서드를 사용하여 현재 활성화된 Google Docs 문서를 가져옵니다.
  var doc = DocumentApp.getActiveDocument();

  // 메서드를 사용하여 문서 본문을 가져옵니다.
  var body = doc.getBody();

  // 메서드를 사용하여 본문에 문단을 추가합니다.
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    body.appendParagraph(row[0] + ': ' + row[1]);
  }

}
