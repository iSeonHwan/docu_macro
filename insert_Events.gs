function insert_Events() {
  //현재 활성화된 시트를 가져옴.
  var sheet = SpreadsheetApp.getActiveSheet();

  //기본 캘린더를 반환함.
  var calendar = CalendarApp.getDefaultCalendar();

  //현재 선택된 셀을 가져옴.
  var selection = sheet.getActiveSelection();
  
  //현재 선택된 셀에 오늘의 날짜를 입력함.
  var now = new Date();
  
  //오늘 날짜의 이벤트를 가져옴.
  var events = calendar.getEventsForDay(now);

  //삽입 위치를 현재 위치 기준으로 설정함.
  var insertRow = selection.getRow();
  var insertColumn = selection.getColumn();

  //에벤트 목록을 스프레드시트에 삽입함.
  for(var i = 0; i < events.length; i++){
    var event = events[i];
    sheet.getRange(insertRow + i, insertColumn).setValue(event.getTitle());
  }

}
