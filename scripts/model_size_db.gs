// 제품모델-사이즈 매핑 테이블 (시트 연동)
// "모델사이즈맵" 시트에 영구 저장, [모델코드, 사이즈, 타입] 3컬럼

/**
 * 모델코드로 사이즈와 타입을 반환. 없으면 다이얼로그로 입력받아 시트에 자동 추가
 * @param {string} model
 * @returns {{size: string, type: string}}
 */
function getSizeByModel(model) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('모델사이즈맵');
  if (!sheet) {
    sheet = ss.insertSheet('모델사이즈맵');
    sheet.appendRow(['모델코드', '모델타입', '사이즈']);
  }
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    map[data[i][0]] = data[i][2]; // 사이즈는 3번째 컬럼
  }
  // 기존에 있으면 반환
  if (map[model]) {
    return map[model];
  }
  // 없으면 prompt로 사이즈 입력받아 시트에 추가
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('신규 모델코드 발견', '모델코드 ' + model + '에 해당하는 사이즈를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var size = response.getResponseText();
    if (size) {
      var lastRow = sheet.getLastRow() + 1;
      sheet.appendRow([model, '', size]);
      // 추가된 행의 B열(모델타입)에 드롭다운 생성
      var typeRange = sheet.getRange(lastRow, 2);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['UF3', 'UF2', 'UV3'], true)
        .setAllowInvalid(false)
        .build();
      typeRange.setDataValidation(rule);
      // 드롭다운 셀로 포커스 이동
      sheet.setActiveRange(typeRange);
      return size;
    }
  }
  return '';
}