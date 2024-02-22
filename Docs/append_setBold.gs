/* 
파일명: append_setBold.gs
설명: 텍스트 삽입 후 굵게 처리하는 코드
작성일: 2024. 2. 22. 
*/

function append_setBold() {
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();


  var text1 = body.appendParagraph('Hello, World!!');
  var text1Length = text1.getText().length;
  text1.editAsText().setBold(0, text1Length-1, true);

}
