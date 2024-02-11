function insert_Today() {
  //현재 활성화된 시트를 가져옴.
  var sheet = SpreadsheetApp.getActiveSheet();

  //기본 캘린더를 반환함.
  var calendar = CalendarApp.getDefaultCalendar();

  //현재 선택된 셀을 가져옴.
  var selection = sheet.getActiveSelection();
  
  //현재 선택된 셀에 오늘의 날짜를 입력함.
  var now = new Date();
  selection.setValue(now);

}
