/* 
파일명: Import_questionnaire.gs
설명: 설문지와 연결된 구글 시트(Sheet)에서 응답자별로 데이터 가져오는 매크로 코드
작성: 2024. 2. 21.
*/

function importSheetData() {
  // 시트 ID 및 범위 입력
  var sheetId = '0000';

  // 시트 열기
  var sheet = SpreadsheetApp.openById(sheetId);

  // 문서 열기
  var doc = DocumentApp.getActiveDocument();

  // 데이터 가져오기
  var data = sheet.getDataRange().getValues();

  //데이터의 행의 길이 수를 구함.
  var dataRow = data.length;
  Logger.log('데이터 행의 수:'+dataRow);
  
  //데이터의 열의 길이 수를 구함.
  var dataCol = data[0].length;
  Logger.log('데이터 열의 수:' + dataCol);

  //항목별로, 학생별로 정보를 가져와 정리한다. 첫 번째 행열은 필요없으므로 생략한다. 그러므로 0부터가 아닌, 1부터 시작하자.
  for(var j = 1; j < dataRow; j++){
    
    //응답자 별로 제목을 넣는다.
    var text1 = doc.appendParagraph( j + '번째 응답');
    text1.setHeading(DocumentApp.ParagraphHeading.TITLE);

    for (var i = 1; i < dataCol; i++){
      try{
        //질문을 넣는다.
        var text2 = doc.appendParagraph(data[0][i]);
        //질문을 개요6으로 설정함.
        text2.setHeading(DocumentApp.ParagraphHeading.HEADING6);

        //응답을 넣는다.
        var text3 = doc.appendParagraph(data[j][i]);
        //응답을 본문으로 설정함.
        text3.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      } catch(error){

      }
    //구분을 위한 줄 바꿈.
    doc.appendParagraph('\n');
    
    }
  //학생별로 페이지를 나눈다.
  doc.appendPageBreak();
  }

}
