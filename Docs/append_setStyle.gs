/* 
파일명: append_setStyle.gs
설명: 텍스트 삽입 후 스타일을 설정하는 코드
작성일: 2024. 2. 22. 
*/

function append_setBold() {
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  //단락을 삽입함.

  var text1 = doc.appendParagraph('Hello, World!!');
  //제목을 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.TITLE); 

  var text1 = doc.appendParagraph('Hello, World!!');
  //부제목을 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.SUBTITLE); 

  var text1 = doc.appendParagraph('Hello, World!!');
  //개요1을 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  //오류가 발생함: doc.setHeading(DocumentApp.ParagraphHeading.HEADING1); => 오류: TypeError: doc.setHeading is not a function

  var text1 = doc.appendParagraph('Hello, World!!');
  //개요2를 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var text1 = doc.appendParagraph('Hello, World!!');
  //개요3을 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.HEADING3);

  var text1 = doc.appendParagraph('Hello, World!!');
  //개요4를 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.HEADING4);

  var text1 = doc.appendParagraph('Hello, World!!');
  //개요5를 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.HEADING5);

  var text1 = doc.appendParagraph('Hello, World!!');
  //개요6를 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.HEADING6);

  var text1 = doc.appendParagraph('Hello, World!!');
  //일반 단락을 설정함.
  text1.setHeading(DocumentApp.ParagraphHeading.NORMAL);

}
