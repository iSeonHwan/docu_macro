/*
definition: 동일 파일 안에서 다른 시트 간 데이터 복사하기.
date: 2024. 3. 7.
 */

function copyDataBetweenSheets() {

  // 원본 시트 및 대상 시트 객체 가져오기
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('원본 시트');
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('대상 시트');

  // '원본 시트' 시트의 특정 범위를 선택합니다.
  const sourceRange = sourceSheet.getRange('A1:A2'); // 원하는 범위로 변경

  // '대상 시트' 시트의 특정 범위에 값을 복사합니다.
  const targetRange = targetSheet.getRange('A1:A2'); // 원하는 범위로 변경
  targetRange.setValues(sourceRange.getValues());

}
