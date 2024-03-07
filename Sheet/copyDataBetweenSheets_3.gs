/*
definition: 동일 파일 안에서 다른 시트 간 데이터 복사하기. ver3
date: 2024. 3. 7.
 */

function copyDataBetweenSheets() {

  // 원본 시트 및 대상 시트 객체 가져오기
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('원본 시트');
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('대상 시트');

  // '원본 시트' 시트의 전체 범위를 선택합니다.
  const sourceRange = sourceSheet.getDataRange();

  // '대상 시트' 시트의 마지막 행 번호를 가져옵니다.
  const lastRow = targetSheet.getLastRow();

  // '대상 시트' 시트의 마지막 행 바로 아래 범위를 선택합니다.
  const targetRange = targetSheet.getRange(lastRow + 1, 1, sourceRange.getWidth(), sourceRange.getHeight());
 

  // '원본 시트' 범위의 값과 서식을 '대상 시트' 범위에 복사합니다.
  sourceRange.copyTo(targetRange);

}
