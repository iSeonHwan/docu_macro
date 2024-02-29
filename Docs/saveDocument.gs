/*
  정의: 문서의 제목을 설정하고 저장하는 코드
  날짜: 2024-2-29.
 */

function saveDocument() {
  // 현재 문서 가져오기
  var doc = DocumentApp.getActiveDocument();

  // 문서 제목 설정
  doc.setName('My Document');

  // 문서 저장
  doc.saveAndClose();

}
