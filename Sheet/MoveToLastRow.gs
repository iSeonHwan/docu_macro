/* 활성화된 영역의 마지막으로 이동하는 코드 */

function moveToLastRow() {
  var sheet = SpreadsheetApp.getActiveSheet();

  //getDataRange 메서드를 사용하여 시트의 모든 데이터 셀 범위를 가져옵니다.
  var range = sheet.getDataRange();

  //getLastColumn 메서드를 사용하여 선택된 범위의 마지막 열 번호를 가져옵니다.
  var lastRow = range.getLastRow();

  //activate 메서드를 사용하여 마지막 열의 첫 번째 셀을 활성화합니다.
  sheet.getRange(lastRow, range.getColumn()).activate();

}
