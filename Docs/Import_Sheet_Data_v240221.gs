/* 구글 시트에서 데이터 가져오기 2024. 2. 21. */

function importSheetData() {
  // 시트 ID 및 범위 입력
  var sheetId = '11bldxPQ7agyw_tqnSS4ywyNkIzQRHCI94s16Vbf3txI';

  // 시트 열기
  var sheet = SpreadsheetApp.openById(sheetId);

  // 문서 열기
  var doc = DocumentApp.getActiveDocument();

  // 데이터 가져오기
  var data = sheet.getDataRange().getValues();

  // 데이터 추가

  // Logger.log()를 사용하여 data 배열 전체를 콘솔에 출력
  //Logger.log(data);

  //data의 길이 출력

  //Logger.log(data.length);
  //Logger.log( data[data.length][0]); //오류 발생: TypeError: Cannot read properties of undefined

  //데이터의 행의 길이 수를 구함.
  var dataRow = data.length;
  Logger.log('데이터 행의 수:'+dataRow);
  
  //데이터의 열의 길이 수를 구함.
  var dataCol = data[0].length;
  Logger.log('데이터 열의 수:' + dataCol);

  /*
  for(var j = 1; j < dataCol; j++){
    for (var i = 0; i < dataCol; i++){
      Logger.log(data[0][i])
      Logger.log(data[j][i])
      Logger.log('\n')
    }
  }
 */

  //항목별로, 학생별로 정보를 가져와 정리한다. 첫 번째 행열은 필요없으므로 생략한다. 그러므로 0부터가 아닌, 1부터 시작하자.
  for(var j = 1; j < dataRow; j++){
    for (var i = 1; i < dataCol; i++){
      try{
        doc.appendParagraph(data[0][i]);
        doc.appendParagraph(data[j][i]);
      } catch(error){

      }
    //구분을 위한 줄 바꿈.
    doc.appendParagraph('\n');
    
    }
  //학생별로 페이지를 나눈다.
  doc.appendPageBreak();
  }

}
