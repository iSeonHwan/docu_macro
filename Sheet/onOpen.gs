/* 스프레드 시트를 실행하자마자 시작하는 동작 */

function onOpen() {
  //특정 시트로 이동한다.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('일지');
  sheet.activate();

 
  //getDataRange 메서드를 사용하여 시트의 모든 데이터 셀 범위를 가져옵니다.
  var range = sheet.getDataRange();

  //getLastColumn 메서드를 사용하여 선택된 범위의 마지막 열 번호를 가져옵니다.
  var lastRow = range.getLastRow();

  //activate 메서드를 사용하여 마지막 열의 첫 번째 셀을 활성화합니다.
  sheet.getRange(lastRow, range.getColumn()).activate();
  
}
