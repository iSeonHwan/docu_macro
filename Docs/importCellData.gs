/* 시트 셀 하나에서 데이터 가져오기 */

function importCellData() {
  // 시트 ID 및 셀 위치 입력
  var sheetId = '0000';
  var cellAddress = 'A1';

  // 시트 열기
  var sheet = SpreadsheetApp.openById(sheetId);

  // 셀 값 가져오기
  var cellValue = sheet.getRange(cellAddress).getValue();

  // 구글 문서 정보 입력
  var docId = '0000';

  // 문서 열기
  var doc = DocumentApp.openById(docId);

  // Gets the document body.
  const body = doc.getBody();

  //var body = DocumentApp.getActiveDocument().getBody();

  // Use editAsText to obtain a single text element containing
  // all the characters in the document.
  var text = body.editAsText();

  // Insert text at the beginning of the document.
  text.insertText(0, cellValue);


  // Make the first half of the document blue.
  //text.setForegroundColor(0, text.getText().length / 2, '#00FFFF');


}
